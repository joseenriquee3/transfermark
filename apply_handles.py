"""
Merge handles.csv (from missing_ig.html) back into players.xlsx.

How to use
----------
1. In your browser, click "Download CSV" on the missing-IG helper page.
   It will save handles.csv to your Downloads folder.
2. Move handles.csv into this same transfermark folder
   (or just download it here directly).
3. Double-click run_apply_handles.bat, OR in a terminal:
       python apply_handles.py

What it does
------------
For every row in handles.csv it finds the matching player in players.xlsx
(by their Transfermarkt URL, which is stable even if the name has accents)
and writes the Instagram URL into the spreadsheet. It also marks the
'IG verified' column with 'manual' so you can tell which links came from
you vs. from Transfermarkt.

If you marked a player "No Instagram exists" in the helper, this script
leaves their Instagram blank but writes 'no_ig' into IG verified so the
next run won't keep asking.
"""

import csv
import os
import sys

import pandas as pd

PLAYERS_FILE = "players.xlsx"
HANDLES_FILE = "handles.csv"


def main():
    if not os.path.exists(PLAYERS_FILE):
        print(f"ERROR: {PLAYERS_FILE} not found in this folder.")
        sys.exit(1)
    if not os.path.exists(HANDLES_FILE):
        print(f"ERROR: {HANDLES_FILE} not found. Download it from missing_ig.html first.")
        sys.exit(1)

    df = pd.read_excel(PLAYERS_FILE)
    if "Transfermarkt" not in df.columns or "Instagram" not in df.columns:
        print("ERROR: players.xlsx is missing 'Transfermarkt' or 'Instagram' columns.")
        sys.exit(1)

    if "IG verified" not in df.columns:
        df["IG verified"] = ""
    if "IG confidence" not in df.columns:
        df["IG confidence"] = ""

    # Read handles.csv into a dict: profile URL -> handle/__NONE__
    handles = {}
    with open(HANDLES_FILE, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            tm = (row.get("Transfermarkt") or "").strip()
            ig = (row.get("Instagram") or "").strip()
            if tm:
                handles[tm] = ig

    if not handles:
        print(f"{HANDLES_FILE} has no rows — nothing to merge.")
        return

    filled, marked_none, unmatched = 0, 0, []
    for tm, val in handles.items():
        mask = df["Transfermarkt"] == tm
        if not mask.any():
            unmatched.append(tm)
            continue
        if val == "__NONE__":
            df.loc[mask, "Instagram"] = ""
            df.loc[mask, "IG confidence"] = "none"
            df.loc[mask, "IG verified"] = "no_ig"
            marked_none += 1
        else:
            df.loc[mask, "Instagram"] = val
            df.loc[mask, "IG confidence"] = "high"
            df.loc[mask, "IG verified"] = "manual"
            filled += 1

    df.to_excel(PLAYERS_FILE, index=False)
    print(f"Done.")
    print(f"  Instagram links filled in:  {filled}")
    print(f"  Marked 'no Instagram':      {marked_none}")
    if unmatched:
        print(f"  Couldn't match (Transfermarkt URL not in players.xlsx): {len(unmatched)}")
        for u in unmatched[:5]:
            print(f"    - {u}")
    print(f"\nUpdated {PLAYERS_FILE}.")


if __name__ == "__main__":
    main()
