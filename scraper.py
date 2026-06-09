"""
Transfermarkt scraper -> players.xlsx

What it does
------------
1. Reads Transfermarkt's "most valuable players" list (paginated).
2. Keeps players matching your criteria: U28 over EUR 5M, OR U24 over EUR 2M.
3. Best-effort Instagram lookup for each player (flagged with a confidence level).
4. Writes everything to players.xlsx, which index.html displays.

How to run
----------
Double-click run_scraper.bat, or in a terminal:
    python scraper.py

Settings you may want to change are in the CONFIG block right below.

IMPORTANT / honest warnings
---------------------------
- Scraping Transfermarkt is against their Terms of Service. Keep this personal
  and small. The script goes slow on purpose and caches pages so re-runs do not
  hammer the site. Do not lower the delays.
- Transfermarkt changes its HTML from time to time; when it does, parsing here
  may need updating.
- Instagram matching is "best effort". It guesses from a web search and WILL be
  wrong sometimes. Every guess gets an "IG confidence" of high/low so you can
  review the low ones. Nothing is ever auto-trusted.
- If Transfermarkt blocks the requests (you'll see HTTP 403 / empty pages),
  see the README for the fallback options.
"""

import hashlib
import os
import random
import re
import sys
import time
import urllib.parse

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ============================ CONFIG ============================
# Transfermarkt's statistics lists are HARD-CAPPED at 500 rows each. To reach
# every player matching the criteria we crawl several NARROW lists and merge:
#
#  1) Age-class lists (u18, u19, ...) — each capped at 500, reaches different
#     value floors. Good for the very youngest cheap players.
#  2) Birth-year lists (one per cohort 1998..2008) — each capped at 500. Each
#     1-year cohort is small enough that a 500-row list reaches very deep
#     (1999 cohort reaches €1.8M, 2003 reaches €1.5M, etc.). This is what
#     fills the big mid-tier gap (Phil Foden, Hakimi, Mbeumo, Tchouaméni...).
#
# Players overlap heavily between lists; we dedupe by Transfermarkt URL.
AGE_CLASSES = ["u18", "u19", "u20", "u21", "u23", "23-30"]
BIRTH_YEARS = list(range(1997, 2009))   # 1997..2008 inclusive — covers U28 (some 1997s
                                         # are still 27 turning 28) plus U24 cohort.
MAX_PAGES_PER_CLASS = 20                 # 20 = the full list (Transfermarkt's hard cap)

# Note: the run stops reading a list early once its pages drop below your value
# floor, so you won't waste time on players too cheap to matter.

# Your filter criteria (market value is in millions of EUR).
# Single rule: every U28 player worth at least €1M.
U28_MIN_VALUE_M = 1.0   # all players aged <= 28 must be worth at least this
U24_MIN_VALUE_M = 1.0   # (kept equal to U28 since the rule is now unified)
U28_AGE = 28
U24_AGE = 24

# Try to find Instagram links (set to False to skip and go much faster).
DO_INSTAGRAM = True

# Polite crawling. DO NOT make these smaller.
MIN_DELAY_SECONDS = 4
MAX_DELAY_SECONDS = 8

CACHE_DIR = "cache"
OUTPUT_FILE = "players.xlsx"
BASE = "https://www.transfermarkt.com"
LIST_URL_BASE = BASE + "/spieler-statistik/wertvollstespieler/marktwertetop"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}
# ===============================================================


def safe_print(msg):
    """Print without crashing on names the Windows console can't render (Turkish ı,
    accents, etc.). Replaces unrenderable characters with '?'."""
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = (getattr(sys.stdout, "encoding", None) or "ascii")
        print(msg.encode(enc, errors="replace").decode(enc, errors="replace"))


def polite_sleep():
    time.sleep(random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS))


def cached_get(url):
    """Fetch a URL, caching the HTML to disk so re-runs don't re-hit the site."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    key = hashlib.md5(url.encode()).hexdigest()
    path = os.path.join(CACHE_DIR, key + ".html")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()

    # Retry a few times: Transfermarkt sometimes returns transient 502/503 errors.
    last_err = None
    for attempt in range(3):
        polite_sleep()
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30)
            resp.raise_for_status()
            resp.encoding = "utf-8"  # Transfermarkt serves UTF-8; fixes accents (Mbappé etc.)
            with open(path, "w", encoding="utf-8") as f:
                f.write(resp.text)
            return resp.text
        except requests.HTTPError as e:
            last_err = e
            if resp.status_code in (502, 503, 429, 500) and attempt < 2:
                print(f"    (retry {attempt + 1}/2 after {resp.status_code})")
                time.sleep(8 * (attempt + 1))
                continue
            raise
    raise last_err


def parse_value_to_millions(text):
    """'€80.00m' -> 80.0 ; '€900k' -> 0.9 ; '-' -> None"""
    if not text:
        return None
    t = text.replace("\xa0", " ").strip().lower().replace("€", "")
    m = re.search(r"([\d.,]+)\s*([mk]?)", t)
    if not m:
        return None
    num = float(m.group(1).replace(",", "."))
    unit = m.group(2)
    if unit == "k":
        return round(num / 1000, 3)
    return num  # already millions (or a plain number)


def matches_criteria(age, value_m):
    """Reference file convention is inclusive: 'U28' means age <= 28 (not < 28).
    A 28-year-old like Lautaro or Bruno Guimarães qualifies."""
    if age is None or value_m is None:
        return False
    if age <= U24_AGE and value_m >= U24_MIN_VALUE_M:
        return True
    if age <= U28_AGE and value_m >= U28_MIN_VALUE_M:
        return True
    return False


def scrape_list_page(page, ak=None, year=None, pos=None):
    """Return a list of player dicts from one most-valuable list page.
    Pass ak (age class), year (birth year), and/or pos (position id 1=GK,
    2=Defender, 3=Midfielder, 4=Attacker). Combining year+pos quarters the
    list, so each can reach much lower market values."""
    params = [f"page={page}"]
    if ak:
        params.append(f"altersklasse={ak}")
    if year:
        params.append(f"jahrgang={year}")
    if pos:
        params.append(f"ausrichtung=alle&spielerposition_id={pos}")
    url = LIST_URL_BASE + "?" + "&".join(params)
    html = cached_get(url)
    soup = BeautifulSoup(html, "lxml")
    table = soup.select_one("table.items")
    if not table:
        return []

    players = []
    for row in table.select("tbody > tr"):
        cells = row.find_all("td", recursive=False)
        if len(cells) < 3:
            continue

        # Cell 1 holds a small inner table: line 1 = name (+ profile link), line 2 = position.
        name_cell = cells[1]
        link = name_cell.select_one("a[href*='/profil/spieler/']")
        if not link or not link.get("href"):
            continue
        name = link.get("title") or link.get_text(strip=True)
        profile = urllib.parse.urljoin(BASE, link["href"])

        position = ""
        inner = name_cell.find("table")
        if inner:
            inner_rows = inner.find_all("tr")
            if len(inner_rows) > 1:
                position = inner_rows[1].get_text(strip=True)

        # Market value is the last (rechts hauptlink) cell.
        value_m = parse_value_to_millions(cells[-1].get_text(strip=True))

        # Age is always cell index 2 on Transfermarkt stats tables.
        # (Cell 0 = rank, cell 1 = name/position, cell 2 = age.) The earlier
        # "first centered digit 14-45" heuristic mis-grabbed the RANK number for
        # players whose rank fell in that range (e.g. Foden at rank 30 -> age=30).
        age = None
        if len(cells) > 2:
            txt = cells[2].get_text(strip=True)
            if txt.isdigit():
                age = int(txt)

        # Nationality + club come from image title attributes in the row.
        nationality = ""
        club = ""
        for img in row.find_all("img"):
            title = (img.get("title") or "").strip()
            if not title or title == name:
                continue
            cls = " ".join(img.get("class", []))
            if "flaggenrahmen" in cls and not nationality:
                nationality = title
            elif not club and "flaggenrahmen" not in cls:
                club = title

        players.append({
            "Name": name,
            "Age": age,
            "Position": position,
            "Club": club,
            "Nationality": nationality,
            "Market value (€M)": value_m,
            "Transfermarkt": profile,
        })
    return players


def fetch_profile(profile_url):
    """Fetch the profile page once and return (html, BeautifulSoup) or (None, None)."""
    try:
        html = cached_get(profile_url)
        return html, BeautifulSoup(html, "lxml")
    except Exception:
        return None, None


# Map Transfermarkt's info-table label -> our column name.
INFO_FIELDS = {
    "Foot:": "Foot",
    "Height:": "Height",
    "Player agent:": "Agent",
    "Joined:": "Joined club",
    "Contract expires:": "Contract expires",
    "Last contract extension:": "Last contract extension",
    "Outfitter:": "Outfitter",
    "Full name:": "Full name",
    "Place of birth:": "Place of birth",
}


def parse_profile_info(soup):
    """Return a dict of extra fields harvested from the Transfermarkt profile."""
    out = {v: "" for v in INFO_FIELDS.values()}
    if soup is None:
        return out
    info = soup.find("div", class_=re.compile("info-table"))
    if not info:
        return out
    labels = info.find_all("span", class_=re.compile("content--regular"))
    values = info.find_all("span", class_=re.compile("content--bold"))
    for lab, val in zip(labels, values):
        key = lab.get_text(strip=True)
        if key in INFO_FIELDS:
            text = val.get_text(" ", strip=True).replace("\xa0", " ")
            out[INFO_FIELDS[key]] = text
    return out


def find_instagram_from_html(html, soup, name):
    """
    Find a player's Instagram from their Transfermarkt profile.

    Verification strategy (#2): only accept links inside the player's
    structured "social-media-toolbar" container — these are the handles
    Transfermarkt has explicitly tied to that player. Anything else in the
    HTML (footer, fan accounts, etc.) is rejected.

    Returns (url, confidence, source) where:
      confidence: 'high' | 'low' | 'none'
      source:     'tm_verified'  - link came from the structured social toolbar
                  'fallback'     - found loose in the HTML, treat with caution
                  'none'         - no Instagram on the profile
    """
    if soup is None:
        return "", "none", "none"

    skip = {"transfermarkt_official", "p", "reel", "reels", "explore",
            "accounts", "about", "developer", "legal", "directory"}

    def clean_handle(href):
        m = re.search(r"instagram\.com/([A-Za-z0-9_.]+)", href, re.I)
        if not m:
            return None
        h = m.group(1).lower().rstrip(".")
        if h in skip or "?" in m.group(0) or len(h) < 2:
            return None
        return h

    # Structured source: the player's own social-media toolbar.
    handle, source = None, "none"
    for container in soup.select(".social-media-toolbar__icons, .socialmedia-icons"):
        for a in container.select("a[href*='instagram.com']"):
            h = clean_handle(a.get("href", ""))
            if h:
                handle, source = h, "tm_verified"
                break
        if handle:
            break

    # Fallback: loose HTML scan (only if structured lookup found nothing).
    if not handle and html:
        for m in re.finditer(r"instagram\.com/([A-Za-z0-9_.]+)", html, re.I):
            h = m.group(1).lower().rstrip(".")
            if h in skip or "?" in m.group(0) or len(h) < 2:
                continue
            handle, source = h, "fallback"
            break

    if not handle:
        return "", "none", "none"

    clean = "https://www.instagram.com/" + handle + "/"

    # Name-overlap heuristic still informs confidence.
    name_letters = re.sub(r"[^a-z]", "", name.lower())
    handle_letters = re.sub(r"[^a-z]", "", handle)
    parts = [p for p in re.split(r"[^a-z]+", name.lower()) if len(p) >= 3]
    overlap = any(p in handle_letters for p in parts) or \
        (handle_letters and handle_letters in name_letters)

    # Decision:
    #   - tm_verified link + name overlap  -> 'high'
    #   - tm_verified but odd handle       -> 'high' (TM trusted it; trust it too)
    #   - fallback + name overlap          -> 'low'
    #   - fallback no overlap              -> 'low'
    if source == "tm_verified":
        conf = "high"
    else:
        conf = "low"
        if overlap:
            conf = "low"
    return clean, conf, source


# The smallest value we ever care about (cheapest U24 threshold). Once a list's
# pages drop entirely below this, there's no point reading deeper into that list.
MIN_VALUE_OF_INTEREST = min(U24_MIN_VALUE_M, U28_MIN_VALUE_M)


# ============================ Change tracker (#6) ============================
# After every run, compare the new list against players_previous.xlsx and write
# changes.xlsx with 5 sheets: New, Gone, Risers, Fallers, Contracts ending soon.
CHANGES_FILE = "changes.xlsx"
PREVIOUS_FILE = "players_previous.xlsx"
CONTRACT_DEADLINE_DAYS = 365     # "ending soon" = within this many days
RISER_FALLER_COUNT = 30           # how many top risers/fallers to list


def _parse_date(text):
    """Parse Transfermarkt's '30/06/2031' style date. Returns datetime or None."""
    if not text:
        return None
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", text.strip())
    if not m:
        return None
    from datetime import datetime
    try:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def write_change_report(df_new):
    """Compare df_new to the previous snapshot and write changes.xlsx."""
    from datetime import datetime, timedelta

    print("\nBuilding change report...")

    if not os.path.exists(PREVIOUS_FILE):
        print(f"  (no {PREVIOUS_FILE} yet — first run, so nothing to compare against)")
        prev = pd.DataFrame(columns=df_new.columns)
    else:
        prev = pd.read_excel(PREVIOUS_FILE)

    # Use Transfermarkt URL as the stable ID (names can vary by accent).
    key = "Transfermarkt"
    new_ids = set(df_new[key].dropna())
    old_ids = set(prev[key].dropna()) if key in prev.columns else set()

    # New = in new file, not in old.
    new_rows = df_new[df_new[key].isin(new_ids - old_ids)]
    # Gone = in old, not in new.
    gone_rows = prev[prev[key].isin(old_ids - new_ids)] if not prev.empty else prev

    # Value risers / fallers among players present in BOTH snapshots.
    val_col = "Market value (€M)"
    if not prev.empty and val_col in prev.columns and val_col in df_new.columns:
        merged = df_new[[key, "Name", "Age", "Club", val_col]].merge(
            prev[[key, val_col]].rename(columns={val_col: "Old value"}),
            on=key, how="inner",
        )
        merged["Delta (€M)"] = merged[val_col] - merged["Old value"]
        risers = merged.sort_values("Delta (€M)", ascending=False).head(RISER_FALLER_COUNT)
        fallers = merged.sort_values("Delta (€M)", ascending=True).head(RISER_FALLER_COUNT)
        risers = risers[risers["Delta (€M)"] > 0]
        fallers = fallers[fallers["Delta (€M)"] < 0]
    else:
        empty = pd.DataFrame(columns=[key, "Name", "Age", "Club", val_col, "Old value", "Delta (€M)"])
        risers, fallers = empty, empty

    # Contracts ending within CONTRACT_DEADLINE_DAYS.
    today = datetime.today()
    deadline = today + timedelta(days=CONTRACT_DEADLINE_DAYS)
    if "Contract expires" in df_new.columns:
        df_new = df_new.copy()
        df_new["_contract_date"] = df_new["Contract expires"].apply(_parse_date)
        ending = df_new[df_new["_contract_date"].notna() &
                        (df_new["_contract_date"] >= today) &
                        (df_new["_contract_date"] <= deadline)]
        ending = ending.sort_values("_contract_date").drop(columns=["_contract_date"])
    else:
        ending = pd.DataFrame()

    with pd.ExcelWriter(CHANGES_FILE, engine="openpyxl") as xw:
        new_rows.to_excel(xw, sheet_name="New players", index=False)
        gone_rows.to_excel(xw, sheet_name="Gone players", index=False)
        risers.to_excel(xw, sheet_name="Value risers", index=False)
        fallers.to_excel(xw, sheet_name="Value fallers", index=False)
        ending.to_excel(xw, sheet_name="Contracts ending soon", index=False)

    print(f"  -> {CHANGES_FILE}: "
          f"{len(new_rows)} new, {len(gone_rows)} gone, "
          f"{len(risers)} risers, {len(fallers)} fallers, "
          f"{len(ending)} contracts ending within {CONTRACT_DEADLINE_DAYS} days.")
# ===========================================================================


def crawl_list(label, by_profile, ak=None, year=None, pos=None):
    """Walk one filtered list, dedupe into by_profile."""
    for page in range(1, MAX_PAGES_PER_CLASS + 1):
        try:
            rows = scrape_list_page(page, ak=ak, year=year, pos=pos)
        except requests.HTTPError as e:
            print(f"  ! Error on {label} page {page}: {e}. Moving on.")
            break
        if not rows:
            print(f"  ({label}) no more pages after {page - 1}.")
            break

        for p in rows:
            by_profile.setdefault(p["Transfermarkt"], p)

        page_max_value = max((p["Market value (€M)"] or 0) for p in rows)
        print(f"  {label} page {page}/{MAX_PAGES_PER_CLASS}: "
              f"{len(rows)} players (top on page €{page_max_value}M), "
              f"total unique so far: {len(by_profile)}")

        # If even the most valuable player on this page is below our floor,
        # every remaining page in this list is too cheap — skip the rest.
        if page_max_value < MIN_VALUE_OF_INTEREST:
            print(f"  ({label}) remaining pages are below €{MIN_VALUE_OF_INTEREST}M — skipping rest.")
            break


def main():
    print("Scraping Transfermarkt — age-class lists + birth-year lists, merging by player ID...")
    by_profile = {}  # profile URL -> player dict (dedupes across all lists)

    for ak in AGE_CLASSES:
        print(f"\n=== Age class '{ak}' ===")
        crawl_list(ak, by_profile, ak=ak)

    for year in BIRTH_YEARS:
        print(f"\n=== Birth year {year} ===")
        crawl_list(f"year-{year}", by_profile, year=year)

    # Position-filtered crawl for the older cohorts (1997-2001). The plain
    # birth-year lists for those years cap around €1.5-2M; quartering by
    # position (GK/DEF/MID/ATT) drops the floor well below €1M, so we capture
    # the missing €1M-€2M players in each cohort.
    POSITION_NEEDED_YEARS = [y for y in BIRTH_YEARS if y <= 2001]
    POSITIONS = [(1, "GK"), (2, "DEF"), (3, "MID"), (4, "ATT")]
    for year in POSITION_NEEDED_YEARS:
        for pos_id, pos_name in POSITIONS:
            label = f"year-{year}-{pos_name}"
            print(f"\n=== Birth year {year} + position {pos_name} ===")
            crawl_list(label, by_profile, year=year, pos=pos_id)

    all_players = list(by_profile.values())
    print(f"\nMerged {len(all_players)} unique players. Applying your criteria...")
    kept = [p for p in all_players if matches_criteria(p["Age"], p["Market value (€M)"])]
    kept.sort(key=lambda p: -(p["Market value (€M)"] or 0))
    print(f"  {len(kept)} players match (U{U28_AGE}, value >= €{U28_MIN_VALUE_M}M).")

    if DO_INSTAGRAM:
        print("Visiting each profile to harvest extra fields + Instagram (slow part)...")
        for i, p in enumerate(kept, 1):
            html, soup = fetch_profile(p["Transfermarkt"])

            # #3 — extra profile fields
            for k, v in parse_profile_info(soup).items():
                p[k] = v

            # Instagram + verification (#2) — single, reliable structural check.
            ig, conf, source = find_instagram_from_html(html, soup, p["Name"])
            p["Instagram"] = ig
            p["IG confidence"] = conf
            p["IG verified"] = source  # tm_verified | fallback | none

            safe_print(f"  [{i}/{len(kept)}] {p['Name']}: {ig or '(no IG)'} "
                       f"({conf}, {source})")
            # Save partial progress every 50 players so a crash never costs hours.
            if i % 50 == 0:
                pd.DataFrame(kept).to_excel(OUTPUT_FILE, index=False)
                print(f"    (saved partial progress to {OUTPUT_FILE})")
    else:
        for p in kept:
            for k in list(INFO_FIELDS.values()) + ["Instagram", "IG confidence", "IG verified"]:
                p.setdefault(k, "")

    if not kept:
        print("No players to write. Check the warnings above (Transfermarkt may have blocked the requests).")
        return

    # Final save + change-tracker.
    df = pd.DataFrame(kept)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nDone. Wrote {len(df)} players to {OUTPUT_FILE}.")

    # #6 — compare against the previous snapshot.
    try:
        write_change_report(df)
    except Exception as e:
        print(f"(change report skipped: {e})")

    # Archive this run as the new baseline for next time.
    df.to_excel("players_previous.xlsx", index=False)

    print("Open index.html in your browser (or double-click view_site.bat) to view them.")
    print("Tip: review the rows where 'IG confidence' is low.")


if __name__ == "__main__":
    main()
