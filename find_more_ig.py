"""
Find more Instagram handles for players that came up empty after the main scrape.

Runs three passes in order. Each pass only works on what the previous passes
couldn't find:

  Pass 1 - Wikidata.       Free, fast, very accurate. Uses the player's
                           Transfermarkt ID to look up Wikidata's curated
                           Instagram username. Most reliable source.

  Pass 2 - DuckDuckGo Lite. Free, slow, mid accuracy. For each remaining
                           missing player, searches "<name> footballer instagram"
                           and extracts the first plausible instagram.com handle.
                           Rate-limited HARD: 8-15s per query.

  Pass 3 - Handle probing. Free, fast-but-unreliable. For each STILL-missing
                           player, generates obvious handle patterns
                           (firstnamelastname, firstname_lastname, etc.) and
                           HEADs Instagram to see which URLs respond as profiles.
                           Marks every hit as 'low' confidence -- you should
                           verify these by hand with missing_ig.html before
                           contacting anyone.

Results are written straight back to players.xlsx. Confidence:
  'high' - Wikidata source or strong handle-name match
  'low'  - DuckDuckGo result or probed handle; eyeball before trusting

Run this AFTER scraper.py has produced players.xlsx.
"""

import csv
import hashlib
import os
import random
import re
import sys
import time
import urllib.parse
from datetime import datetime

import pandas as pd
import requests

PLAYERS_FILE = "players.xlsx"
CACHE_DIR = "cache_ig"            # separate cache so the main scraper's cache stays clean
REPORT_FILE = "ig_recovery_report.csv"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
WIKIDATA_UA = "transfermark-personal-tool/1.0 (contact: darksteamlass@gmail.com)"

WIKIDATA_BATCH_SIZE = 40           # how many TM IDs per SPARQL query

# Polite delays for the search and IG probing passes.
DDG_MIN_DELAY = 8
DDG_MAX_DELAY = 15
IG_PROBE_MIN_DELAY = 3
IG_PROBE_MAX_DELAY = 6


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------
def safe_print(msg):
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = (getattr(sys.stdout, "encoding", None) or "ascii")
        print(msg.encode(enc, errors="replace").decode(enc, errors="replace"))


def cached_get(url, ua=USER_AGENT, delay=(4, 8), extra_headers=None):
    """Fetch with delay + on-disk cache. Returns text or raises."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    key = hashlib.md5(url.encode()).hexdigest()
    path = os.path.join(CACHE_DIR, key + ".html")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    time.sleep(random.uniform(*delay))
    headers = {"User-Agent": ua, "Accept-Language": "en-US,en;q=0.9"}
    if extra_headers:
        headers.update(extra_headers)
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    with open(path, "w", encoding="utf-8") as f:
        f.write(resp.text)
    return resp.text


def extract_tm_id(profile_url):
    """https://www.transfermarkt.com/.../spieler/937958 -> '937958'"""
    if not isinstance(profile_url, str):
        return None
    m = re.search(r"/spieler/(\d+)", profile_url)
    return m.group(1) if m else None


def normalize_ig_handle(raw):
    """'@LamineYamal' / 'https://instagram.com/lamineyamal/' -> 'lamineyamal'."""
    if not raw:
        return None
    s = str(raw).strip()
    m = re.search(r"instagram\.com/([A-Za-z0-9_.]+)", s)
    if m:
        s = m.group(1)
    s = s.lstrip("@").rstrip("/").lower()
    if not re.fullmatch(r"[a-z0-9_.]{2,30}", s):
        return None
    if s in {"p", "reel", "reels", "explore", "accounts", "about",
             "transfermarkt_official", "instagram"}:
        return None
    return s


def to_url(handle):
    return f"https://www.instagram.com/{handle}/"


# ---------------------------------------------------------------------------
# Pass 1 -- Wikidata
# ---------------------------------------------------------------------------
def pass1_wikidata(df_missing):
    """Batch-query Wikidata for every missing player's TM ID at once."""
    print(f"\n=== Pass 1 - Wikidata ({len(df_missing)} players) ===")
    found = {}  # TM url -> handle
    tm_id_to_url = {}
    for _, row in df_missing.iterrows():
        tm_id = extract_tm_id(row["Transfermarkt"])
        if tm_id:
            tm_id_to_url[tm_id] = row["Transfermarkt"]

    ids = list(tm_id_to_url.keys())
    print(f"  {len(ids)} players have a valid TM ID. Querying Wikidata in batches of {WIKIDATA_BATCH_SIZE}...")

    endpoint = "https://query.wikidata.org/sparql"
    for i in range(0, len(ids), WIKIDATA_BATCH_SIZE):
        batch = ids[i:i + WIKIDATA_BATCH_SIZE]
        values = " ".join(f'"{x}"' for x in batch)
        query = (
            "SELECT ?tm ?ig WHERE {\n"
            f"  VALUES ?tm {{ {values} }}\n"
            "  ?player wdt:P2446 ?tm .\n"
            "  OPTIONAL { ?player wdt:P2003 ?ig . }\n"
            "}\n"
        )
        try:
            r = requests.get(
                endpoint,
                params={"query": query, "format": "json"},
                headers={"User-Agent": WIKIDATA_UA,
                         "Accept": "application/sparql-results+json"},
                timeout=60,
            )
            r.raise_for_status()
            batch_hits = 0
            for binding in r.json()["results"]["bindings"]:
                tm = binding["tm"]["value"]
                ig_raw = binding.get("ig", {}).get("value")
                handle = normalize_ig_handle(ig_raw) if ig_raw else None
                if handle and tm in tm_id_to_url:
                    found[tm_id_to_url[tm]] = handle
                    batch_hits += 1
            done = min(i + WIKIDATA_BATCH_SIZE, len(ids))
            print(f"  batch {i // WIKIDATA_BATCH_SIZE + 1}: queried {len(batch)} TM IDs, {batch_hits} new handles. Total so far: {len(found)} / {done} queried")
            time.sleep(1.5)  # be polite to Wikidata
        except Exception as e:
            print(f"  batch {i // WIKIDATA_BATCH_SIZE + 1} failed: {e}")
            time.sleep(5)

    print(f"  Pass 1 done: {len(found)} new IG handles via Wikidata.")
    return found  # TM url -> handle


# ---------------------------------------------------------------------------
# Pass 2 -- DuckDuckGo Lite search
# ---------------------------------------------------------------------------
def pass2_ddg(df_missing):
    """Search '<name> footballer instagram' on DuckDuckGo Lite, extract first hit."""
    print(f"\n=== Pass 2 - DuckDuckGo Lite ({len(df_missing)} players, slow) ===")
    found = {}
    skip_handles = {"transfermarkt_official", "p", "reel", "explore",
                    "accounts", "instagram"}
    handle_re = re.compile(r"instagram\.com/([A-Za-z0-9_.]+)", re.I)

    for n, (_, row) in enumerate(df_missing.iterrows(), 1):
        name = row["Name"]
        club = str(row.get("Club", "") or "")
        query = f'"{name}" footballer instagram'
        url = "https://lite.duckduckgo.com/lite/?q=" + urllib.parse.quote(query)

        try:
            html = cached_get(url, delay=(DDG_MIN_DELAY, DDG_MAX_DELAY))
        except Exception as e:
            safe_print(f"  [{n}/{len(df_missing)}] {name}: search failed ({e})")
            continue

        handle = None
        for m in handle_re.finditer(html):
            h = m.group(1).lower().rstrip(".")
            if h in skip_handles or "?" in m.group(0) or len(h) < 2:
                continue
            handle = h
            break

        if handle:
            found[row["Transfermarkt"]] = handle
            safe_print(f"  [{n}/{len(df_missing)}] {name} -> {handle}")
        else:
            safe_print(f"  [{n}/{len(df_missing)}] {name}: no hit")

        # Save partial progress every 100 in case we get interrupted.
        if n % 100 == 0:
            _save_intermediate(found, suffix="pass2")

    print(f"  Pass 2 done: {len(found)} candidate handles via DuckDuckGo.")
    return found


# ---------------------------------------------------------------------------
# Pass 3 -- Handle pattern probing
# ---------------------------------------------------------------------------
def generate_candidate_handles(name):
    """Return up to ~8 plausible handles for a player's name."""
    # Strip accents by encoding to ASCII.
    import unicodedata
    nfkd = unicodedata.normalize("NFKD", name)
    ascii_name = "".join(c for c in nfkd if not unicodedata.combining(c))
    ascii_name = ascii_name.lower()
    parts = re.split(r"[^a-z0-9]+", ascii_name)
    parts = [p for p in parts if p]
    if len(parts) < 2:
        return [parts[0]] if parts else []

    first = parts[0]
    last = parts[-1]
    candidates = [
        f"{first}{last}",
        f"{first}.{last}",
        f"{first}_{last}",
        f"{first[0]}{last}",
        f"{first[0]}.{last}",
        f"{last}{first}",
        f"{last}.{first}",
        f"{first}{last}10",       # common: shirt number suffix
    ]
    # Deduplicate while preserving order.
    seen, out = set(), []
    for c in candidates:
        if c not in seen and 2 <= len(c) <= 30:
            seen.add(c)
            out.append(c)
    return out


def probe_handle_exists(handle):
    """HEAD instagram.com/<handle>/ to see if it's a real profile.
    Returns True if it likely exists, False otherwise.

    Instagram returns 200 for real profiles and 404 for nonexistent.
    Login walls return 200 too -- so a 200 means 'maybe', not 'definitely'.
    We treat 200 as exists; 404 as doesn't."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_path = os.path.join(CACHE_DIR, f"probe_{handle}.txt")
    if os.path.exists(cache_path):
        return open(cache_path).read().strip() == "exists"
    time.sleep(random.uniform(IG_PROBE_MIN_DELAY, IG_PROBE_MAX_DELAY))
    try:
        # GET is more reliable than HEAD against Instagram's CDN.
        r = requests.get(to_url(handle), headers={"User-Agent": USER_AGENT}, timeout=20,
                         allow_redirects=True)
        # IG sometimes redirects nonexistent handles to a login page; the URL
        # changing to '/accounts/login' is a strong "doesn't exist" signal.
        exists = (r.status_code == 200
                  and "accounts/login" not in r.url
                  and "page not found" not in r.text.lower()[:5000])
    except Exception:
        exists = False
    with open(cache_path, "w") as f:
        f.write("exists" if exists else "no")
    return exists


def pass3_probe(df_missing):
    """For each missing player, try obvious handle patterns and keep ones that exist."""
    print(f"\n=== Pass 3 - Handle probing ({len(df_missing)} players, mid speed) ===")
    found = {}
    for n, (_, row) in enumerate(df_missing.iterrows(), 1):
        name = row["Name"]
        candidates = generate_candidate_handles(name)
        hit = None
        for h in candidates:
            try:
                if probe_handle_exists(h):
                    hit = h
                    break
            except Exception:
                continue
        if hit:
            found[row["Transfermarkt"]] = hit
            safe_print(f"  [{n}/{len(df_missing)}] {name} -> {hit} (low conf, verify!)")
        else:
            safe_print(f"  [{n}/{len(df_missing)}] {name}: none of {len(candidates)} patterns matched")
        if n % 100 == 0:
            _save_intermediate(found, suffix="pass3")
    print(f"  Pass 3 done: {len(found)} possible handles via pattern probing.")
    return found


# ---------------------------------------------------------------------------
# Save partial progress + final merge
# ---------------------------------------------------------------------------
def _save_intermediate(found, suffix):
    path = f"ig_partial_{suffix}.csv"
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Transfermarkt", "Handle"])
        for tm, h in found.items():
            w.writerow([tm, h])


def apply_to_xlsx(df, results_by_pass):
    """Write found handles back into the dataframe and save to xlsx."""
    # Track what came from where for the report
    sources = {}   # TM url -> pass name
    for pass_name, found in results_by_pass.items():
        for tm, _ in found.items():
            sources.setdefault(tm, pass_name)

    for pass_name in ["wikidata", "duckduckgo", "probe"]:
        found = results_by_pass.get(pass_name, {})
        confidence = "high" if pass_name == "wikidata" else "low"
        verified = {"wikidata": "wikidata",
                    "duckduckgo": "ddg_search",
                    "probe": "pattern_probe"}[pass_name]
        for tm, handle in found.items():
            if sources.get(tm) != pass_name:
                continue  # earlier pass already claimed this row
            mask = df["Transfermarkt"] == tm
            if not mask.any():
                continue
            df.loc[mask, "Instagram"] = to_url(handle)
            df.loc[mask, "IG confidence"] = confidence
            df.loc[mask, "IG verified"] = verified

    df.to_excel(PLAYERS_FILE, index=False)
    return sources


def write_report(df, sources, before_missing):
    after_missing = df[df["Instagram"].fillna("").astype(str).str.strip() == ""]
    print()
    print("=" * 60)
    print("RECOVERY REPORT")
    print("=" * 60)
    print(f"Players missing IG before:  {before_missing}")
    print(f"Players missing IG after:   {len(after_missing)}")
    print(f"Recovered:                  {before_missing - len(after_missing)}")
    print()
    by_source = {}
    for tm, src in sources.items():
        by_source[src] = by_source.get(src, 0) + 1
    print("By pass:")
    for src in ["wikidata", "duckduckgo", "probe"]:
        print(f"  {src:11}: {by_source.get(src, 0)}")
    print()
    pct = 100 * (1 - len(after_missing) / len(df))
    print(f"Total IG coverage now: {pct:.1f}% of {len(df):,} players")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    """
    Wikidata-only recovery.

    Earlier design called for 3 passes (Wikidata + DuckDuckGo + handle probing)
    but live testing showed:
      - DuckDuckGo Lite now serves an anti-bot 'anomaly' page (no results)
      - Instagram's login wall returns 200 OK for *every* handle, so probing
        cannot distinguish real handles from nonsense (it produces false hits
        for any spelling).
    So we keep only Pass 1 (Wikidata), which is the one that actually works.
    Anything Wikidata can't find should go through missing_ig.html by hand.
    """
    if not os.path.exists(PLAYERS_FILE):
        print(f"ERROR: {PLAYERS_FILE} not found.")
        sys.exit(1)

    df = pd.read_excel(PLAYERS_FILE)
    missing = df[df["Instagram"].fillna("").astype(str).str.strip() == ""].copy()
    print(f"Loaded {len(df):,} players. {len(missing):,} are missing Instagram.\n")
    before = len(missing)

    results = {"wikidata": pass1_wikidata(missing)}
    sources = apply_to_xlsx(df, results)
    write_report(df, sources, before)


if __name__ == "__main__":
    main()
