# Redirect Map Automation — Step-by-Step Guide

This covers setting up your machine once, then the repeatable steps for
every new client migration using the point-and-click **GUI**
(`redirect_map_gui.py`). Terminal commands are also documented near the
end for anyone who prefers them or needs a flag the GUI doesn't expose.

---

## Part 1 — One-Time Setup

### Step 1: Install Python (disregard if already installed)

Open **Command Prompt** or **PowerShell** and run:

```
winget install Python.Python.3.13
```

Close and reopen your terminal after install finishes (so it picks up the
new PATH), then confirm it worked:

```
python --version
pip --version
```

You should see something like `Python 3.13.x` and a pip version number. If
`python` isn't recognized, restart your PC once — winget sometimes needs a
fresh terminal session to register PATH changes.

### Step 2: Install Git (disregard if already installed)

Still in the terminal:

```
winget install --id Git.Git -e
```

Close and reopen your terminal after it finishes, then confirm:

```
git --version
```

### Step 3: Install the required Python packages

```
pip install requests beautifulsoup4 lxml openpyxl
```

Optional extra polish for the GUI (not required — it's picked up
automatically if present, and the GUI looks fine without it):

```
pip install sv-ttk
```

Steps 1–3 only need to be done once per machine (or again if you reinstall
Python).

### Step 4: Clone the scripts repo

Pick a folder to keep it in — for example your `Downloads\Jade` folder —
then run:

```
cd C:\Users\Production\Downloads\Jade
git clone https://github.com/jaadddeee/301-automate.git
```

This creates a `301-automate` folder containing three files, all of which
need to stay side by side:
- `redirect_map_gui.py` — the point-and-click app you'll actually open
  (see Part 2)
- `sitemap_to_redirect_map.py` — builds the map from the old site (the GUI
  runs this for you)
- `match_new_site.py` — matches in the new site (the GUI runs this too;
  imports functions directly from the file above)

**Getting future updates:** whenever the scripts get updated in the repo,
just pull the latest version instead of re-cloning:

```
cd C:\Users\Production\Downloads\Jade\301-automate
git pull
```

---

## Part 2 — Running It from the GUI

### Opening it

Double-click `redirect_map_gui.py` in the `301-automate` folder, or run:

```
python redirect_map_gui.py
```

**Tip:** to skip the black console window, right-click the file → Create
shortcut → right-click the shortcut → Properties → change the target so it
starts with `pythonw.exe` instead of `python.exe`. Pin that shortcut
wherever's convenient for one-click access.

### Step ① (301-RW) — Build the redirect map from the OLD site

1. Paste the old (live) site's URL into the **Old site URL** box.
2. The **Output file** name auto-fills based on the domain (e.g.
   `CompName301RW.xlsx`) — edit it directly, or click **Browse…** to pick a
   location.
3. Click **Build Redirect Map from Old Site**.

Watch the **Progress log** panel while it runs. Behind the scenes it:
- Looks for a sitemap first (via `robots.txt`, then common paths like
  `sitemap.xml`).
- Falls back to crawling the site by following links from the homepage if
  no sitemap exists — common on Proweaver/custom-themed sites. You'll see
  `Falling back to crawling the site...` in the log; that's expected, not
  an error.
- Resolves each crawled page against its own `<link rel="canonical">` tag,
  so a hierarchical WordPress page reachable at more than one path (a flat
  nav link to a page that's actually a child of another page) gets
  recorded at its correct full path instead of a flat shortcut, and
  duplicate pages reachable via two different links collapse to one row.
- Checks the WordPress REST API for any pages with no incoming links
  anywhere on the site (orphaned pages a crawl alone would miss).
- Skips individual blog posts, `.php`/`.pdf` files, category/tag/author
  archive pages, and feeds automatically.
- **If a site's `<title>` tag is broken and identical on every page** (a
  real templating bug some sites have), it falls back to that page's
  on-page heading instead, so you don't end up with a spreadsheet full of
  duplicate titles. Look for a `Duplicate title detected for ... -> using
  on-page heading "..." instead` line when this kicks in — worth a glance
  to confirm the substituted title looks right.

When it finishes, the log turns green with `✔ Done.` and the status bar
updates. **Skim the log** before moving on — check how many pages it
found, whether it fell back to crawling, and whether the REST API turned
up any extra orphaned pages.

### Step ② (301-3D) — Match in the NEW site

Once the new site is built (staging or live):

1. Paste the new site's URL into the **New site URL** box.
2. The **Workbook to update** field reuses whatever's in Step ①'s output
   box — leave it as-is, or **Browse…** to point at a different `.xlsx`.
3. Click **Match New Site into Workbook**.

This crawls the new site the same way as Step ① (sitemap first, crawl
fallback, canonical-URL resolution, same broken-title fallback), then for
every new-site page:
- **Exact title match** against Column A → fills that row's Column C with
  the new page's slug.
- **No match** → appends a brand-new row with the title in A and the slug
  in C (Column B stays blank, since there's no old-site equivalent).

It also updates the hyperlink in cell A1 to point at the new site, and
overwrites the same workbook by default.

**Before running this step, close the workbook in Excel if you have it
open** — see Troubleshooting below.

### Advanced options

Each step's card has a collapsed **Advanced options ▸** section — click to
expand it. These map directly to command-line flags, just tucked away
until you need them:

| Option | What it does |
|---|---|
| Max pages to crawl | Cap how many pages the crawl fallback will visit (default 300) |
| Don't fall back to crawling | Fail instead of crawling if no sitemap is found |
| Skip WordPress REST API check | Skip the orphan-page check (Step ① only) |
| Extra URLs (one per line) | Manually add pages the crawler might miss, e.g. truly orphaned pages (Step ① only) |

### Step ③ — Review the workbook by hand

Because matching is exact-title-only, open the spreadsheet afterward and
check for:
- **Rows with a filled Column B but still-blank Column C** — an old page
  that didn't find a match on the new site. Either the new site doesn't
  have that page (needs a redirect decision), or its title changed enough
  that it wasn't recognized — search the new site manually and fill in
  Column C.
- **Appended rows near the bottom with blank Column B** — new-site pages
  that didn't match anything old. Confirm these are genuinely new pages and
  not a renamed old page that should've matched an existing row instead.

### Other GUI buttons

- **Open Output Folder** jumps straight to wherever the workbook was
  saved.
- **Clear** (top-right of the Progress log) clears the log panel without
  affecting anything already saved.

---

## Part 3 — Running It from the Terminal (optional)

The GUI runs these same two scripts for you — this section is for anyone
who prefers typing commands directly, or needs a flag combination faster
than clicking through Advanced options.

### Step 1: Open a terminal in the cloned repo folder

In File Explorer, open the `301-automate` folder (from the `git clone` step
in Part 1), click the address bar, type `cmd` (or `powershell`), and hit
Enter. This opens a terminal already pointed at that folder.

### Step 2: Build the redirect map from the OLD site

```
python sitemap_to_redirect_map.py https://www.oldsite.com/ -o CompName301RW.xlsx
```

Replace `https://www.oldsite.com/` with the live client site being
replaced, and `CompName301RW.xlsx` with whatever naming convention you're
using for that client. Behavior is identical to Step ① in the GUI (see
Part 2 above for what happens under the hood).

### Step 3: Match in the NEW site

```
python match_new_site.py CompName301RW.xlsx https://www.newsite.com/
```

Behavior is identical to Step ② in the GUI. This overwrites the same
`.xlsx` file by default. Add `-o` if you want to save the merged result
somewhere else instead:

```
python match_new_site.py CompName301RW.xlsx https://www.newsite.com/ -o CompName301RW-final.xlsx
```

**Before running this step, close the workbook in Excel if you have it
open** — see Troubleshooting below.

### Useful flags (either script)

| Flag | What it does |
|---|---|
| `-o <path>` | Custom output file path |
| `--max-pages <N>` | Cap how many pages the crawl fallback will visit (default 300) |
| `--no-crawl` | Fail instead of crawling if no sitemap is found |
| `--no-rest-api` | Skip the WordPress REST API orphan-page check (`sitemap_to_redirect_map.py` only) |
| `--extra-urls url1 url2 ...` | Manually add pages the crawler might miss, e.g. truly orphaned pages (`sitemap_to_redirect_map.py` only) |

Example with extras:

```
python sitemap_to_redirect_map.py https://www.oldsite.com/ -o CompName301RW.xlsx --extra-urls https://www.oldsite.com/hidden-page https://www.oldsite.com/another-page
```

---

## Quick Troubleshooting

- **`'python' is not recognized...`** — Python isn't on PATH yet. Reopen
  the terminal, or restart the PC, then try again.
- **`bs4.exceptions.FeatureNotFound: Couldn't find a tree builder with the
  features you requested: xml`** — the `lxml` package is missing. Run
  `pip install lxml` and try again.
- **`PermissionError` / a message about the file being open** — the
  workbook is currently open in Excel (or another program) and Windows
  won't let the script overwrite it. **Close the workbook in Excel, then
  run it again** (click the button again in the GUI, or re-run the
  command in the terminal).
- **`Couldn't locate a sitemap for ...`** — expected on many
  Proweaver-style sites; it automatically falls back to crawling. No action
  needed unless you enabled "Don't fall back to crawling."
- **Fewer pages found than expected** — try raising "Max pages to crawl,"
  or check if some pages are genuinely orphaned (no internal links pointing
  to them) and add them via "Extra URLs."
- **A page's title looks wrong (shows the URL slug instead)** — that page
  timed out or failed to load when the script fetched it; re-run, or check
  the page loads fine in a browser.
- **Every page has the same generic title in the spreadsheet** — some
  sites genuinely serve an identical `<title>` tag on every page (a
  templating bug on their end, not a bug in the script). It's detected
  automatically and substituted with each page's on-page heading instead —
  look for `Duplicate title detected for ...` lines in the log confirming
  it happened, and double check those specific rows.
- **`git pull` says your local changes would be overwritten** — you (or
  someone) edited a script directly instead of pulling updates. Rename
  your edited copy, run `git pull`, then reapply whatever you'd changed.
- **GUI shows "Scripts not found"** — `redirect_map_gui.py` needs to sit in
  the same folder as `sitemap_to_redirect_map.py` and `match_new_site.py`.
  Move whichever one is out of place.