# Player Data — Transfermarkt scraper + website

This is a clone of the site you found, but more complete. It does two things:

1. **Scrapes players from Transfermarkt** (the most-valuable-players list),
   keeps the young/valuable ones, and finds each player's **real Instagram**
   straight from their Transfermarkt profile.
2. **Shows them on a web page** as a sortable, searchable table with buttons
   that open the players' Instagram pages.

You do **not** need to know any coding to use it. Just double-click the files
below.

---

## What's in this folder

| File | What it's for |
|------|----------------|
| `index.html` | The website. Open it through `view_site.bat`. |
| `view_site.bat` | Double-click to open the website (starts a tiny local server). |
| `players.xlsx` | The spreadsheet of players the website reads. |
| `players_previous.xlsx` | Snapshot from the previous run, used by the change-tracker. |
| `changes.xlsx` | New / gone / risers / fallers / contracts-ending-soon report. |
| `run_scraper.bat` | Double-click to refresh players from Transfermarkt. |
| `scraper.py` | The scraper program (you don't need to open this). |
| `missing_ig.html` | Helper page to manually fill in missing Instagram links. |
| `fill_missing_ig.bat` | Double-click to open the missing-IG helper. |
| `apply_handles.py` | Merges the helper's `handles.csv` back into `players.xlsx`. |
| `run_apply_handles.bat` | Double-click to apply the handles you collected. |
| `requirements.txt` | List of helper libraries (installed once, see below). |
| `cache/` | Saved copies of pages so re-runs don't re-download. Safe to delete (re-runs will be slow once). |

---

## First-time setup (do this once)

1. Make sure **Python** is installed. It already is on this machine
   (Python 3.13). If you ever move to a new PC, get it from
   <https://www.python.org/downloads/> and tick *"Add Python to PATH"* during
   install.
2. Install the helper libraries. Open this folder, click the address bar at the
   top of the window, type `cmd`, press Enter, then paste this and press Enter:

   ```
   python -m pip install -r requirements.txt
   ```

That's it. You only do this once.

---

## How to use it

### Just look at the data
Double-click **`view_site.bat`** — opens the website with the current
`players.xlsx` in your browser. Close the black window when you're done.

### Refresh the data from Transfermarkt
Double-click **`run_scraper.bat`**.

- It goes slowly on purpose so Transfermarkt doesn't block it. First run
  takes about an hour or two; re-runs are much faster thanks to the cache.
- When it finishes it writes `players.xlsx` + `changes.xlsx` (the diff
  against the previous run) and opens the website automatically.

### Fill in missing Instagram links by hand
Double-click **`fill_missing_ig.bat`**. The helper page shows every player
without an Instagram, one at a time, with a one-click Google search prefilled
for that player's name. Paste the handle, hit Save, repeat. When done, click
**Download CSV**, drop `handles.csv` into this folder, then double-click
**`run_apply_handles.bat`** to merge them in.

---

## Using the website

- **Search box** — type a name or club to filter the table.
- **Click any column heading** — sorts by that column (click again to reverse).
- **Min / Max value + "Open Instagram (in range)"** — opens the Instagram pages
  of every player whose market value is between those two numbers.
  - First pick the right column in the **"Value column"** dropdown
    (it should already be set to *Market value*).
- **"Open all Instagram"** — opens the Instagram page of every player shown.
- The first time you click an "Open Instagram" button, your browser may block
  the pop-up tabs. Click *"Always allow pop-ups from this site"* and try again.

---

## About the Instagram column

The Instagram links come from each player's **own Transfermarkt profile**, so
they're usually correct. Each row has an **"IG confidence"** column:

- `high` — the handle clearly matches the player's name. Almost always right.
- `low` — worth a quick double-check.
- `none` — Transfermarkt had no Instagram listed for that player.

There's a blank **"IG verified"** column for you to mark (e.g. type "yes")
after you've eyeballed the `low` ones.

---

## Changing the criteria

If you want different players, open `scraper.py` in Notepad and edit the
numbers near the top, in the **CONFIG** section:

- `MAX_PAGES` — how many pages to read (25 players per page). More = more
  players but slower.
- `U28_MIN_VALUE_M` / `U24_MIN_VALUE_M` — the value thresholds (in € millions).
- `DO_INSTAGRAM` — set to `False` to skip Instagram and run much faster.

Then double-click `run_scraper.bat` again.

---

## Putting it online with GitHub Pages (free, works on your phone)

This folder is ready to publish. Once it's online your site will be at
`https://YOURUSERNAME.github.io/REPONAME/` — open it on any phone or PC.

There's a `.gitignore` already set up so the cache folder (700+ MB) and log
files stay on your PC and are NOT uploaded.

### One-time setup (about 5 minutes)

1. **Install GitHub Desktop** (free, official):
   <https://desktop.github.com/>. Sign in with the GitHub account you
   already use in Chrome.
2. **Create a new repository on GitHub.com**:
   - Go to <https://github.com/new>
   - Repository name: e.g. `transfermark` (this becomes part of your URL)
   - Visibility: **Public** (free Pages requires public)
   - Do NOT tick "Add a README" — this folder already has one
   - Click **Create repository**
3. **Link this folder to that empty repo** in GitHub Desktop:
   - File → Add Local Repository → pick this `transfermark` folder
   - GitHub Desktop will say "this isn't a Git repository — would you like
     to create one?" → click **create a repository**
   - Then click **Publish repository** and pick the empty repo you just made
     on GitHub.com (untick "Keep this code private")
4. **Enable GitHub Pages**:
   - On GitHub.com, open your new repo
   - Settings → Pages → under "Source", pick branch **main** and folder **/ (root)**
   - Save. Within a minute, your site is live at the URL shown.

### Whenever you want to update the data

1. Double-click `run_scraper.bat` to fetch fresh players.
2. Open GitHub Desktop. It'll show the files that changed (probably just
   `players.xlsx` and `changes.xlsx`).
3. Type a short message in the bottom-left ("update players May 30" or
   whatever) and click **Commit to main**.
4. Click **Push origin** (top of window).
5. Within ~1 minute the live site shows the new data. Refresh your phone.

### What gets uploaded vs stays on your PC

| Uploaded to GitHub | Stays on your PC only |
|--------------------|------------------------|
| `index.html`, `missing_ig.html` | `cache/` (huge, regenerable) |
| `players.xlsx`, `changes.xlsx` | `cache_ig/` |
| `scraper.py`, `apply_handles.py`, `find_more_ig.py` | `scrape_log.txt` |
| `README.md`, `requirements.txt` | `players_previous.xlsx` (needed for change-tracking) |
| The `.bat` launchers | `handles.csv` if you make one |

---

## Honest warnings

- **Transfermarkt does not officially allow scraping.** Keep this small and
  personal. The scraper deliberately goes slow and saves pages to `cache/` so
  it doesn't hammer the site. Please don't lower the delays in `scraper.py`.
- **Transfermarkt occasionally changes its website.** If the scraper one day
  prints "No players parsed" or stops finding data, the page layout probably
  changed and the parsing code needs a small update.
- If this ever becomes public or commercial, the proper solution is a paid,
  licensed football-data service rather than scraping.
