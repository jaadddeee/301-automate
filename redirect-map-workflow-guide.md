# Redirect Map Automation — Step-by-Step Guide

This covers setting up your machine once, then the repeatable steps for
every new client migration.

---

## Part 1 — One-Time Setup

### Step 1: Install Python

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

### Step 2: Install Git

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
pip install requests beautifulsoup4 openpyxl
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

This creates a `301-automate` folder containing `sitemap_to_redirect_map.py`
and `match_new_site.py` together (they need to stay side by side, since
`match_new_site.py` imports functions directly from the other file).

**Getting future updates:** whenever the scripts get updated in the repo,
just pull the latest version instead of re-cloning:

```
cd C:\Users\Production\Downloads\Jade\301-automate
git pull
```

---

## Part 2 — Running It on an Actual Migration

### Step 1: Open a terminal in the cloned repo folder

In File Explorer, open the `301-automate` folder (from the `git clone` step
above), click the address bar, type `cmd` (or `powershell`), and hit Enter.
This opens a terminal already pointed at that folder.

### Step 2: Build the redirect map from the OLD site

```
python sitemap_to_redirect_map.py https://www.oldsite.com/ -o CompName301RW.xlsx
```

Replace `https://www.oldsite.com/` with the live client site being
replaced, and `CompName301RW.xlsx` with whatever naming convention you're
using for that client.

What happens:
- It looks for a sitemap first (via `robots.txt`, then common paths like
  `sitemap.xml`).
- If no sitemap exists — common on Proweaver/custom-themed sites — it
  automatically falls back to crawling the site by following links from the
  homepage. You'll see `Falling back to crawling the site...` in the
  output; that's expected, not an error.
- It also checks the WordPress REST API for any pages with no incoming
  links anywhere on the site (orphaned pages a crawl alone would miss).
- It skips individual blog posts, `.php`/`.pdf` files, category/tag/author
  archive pages, and feeds automatically.
- It writes the `.xlsx` with Column A = page title, Column B = old URL,
  Column C = blank (to be filled in Step 3).

**Check the console output** before moving on — skim for how many pages it
found, whether it fell back to crawling, and whether the REST API turned up
any extra orphaned pages.

### Step 3: Match in the NEW site

Once the new site is built (staging or live):

```
python match_new_site.py CompName301RW.xlsx https://www.newsite.com/
```

What happens:
- It crawls the new site the same way (sitemap first, crawl fallback).
- For every new-site page, it compares the page title against Column A:
  - **Exact title match** → fills that row's Column C with the new page's
    slug.
  - **No match** → appends a brand-new row with the title in A and the slug
    in C (Column B stays blank, since there's no old-site equivalent).
- It updates the hyperlink in A1 to point at the new site.

This overwrites the same `.xlsx` file by default. Add `-o` if you want to
save the merged result somewhere else instead:

```
python match_new_site.py CompName301RW.xlsx https://www.newsite.com/ -o CompName301RW-final.xlsx
```

### Step 4: Review the workbook by hand

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

---

## Useful Flags (either script)

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
- **`Couldn't locate a sitemap for ...`** — expected on many
  Proweaver-style sites; it automatically falls back to crawling. No action
  needed unless you passed `--no-crawl`.
- **Fewer pages found than expected** — try raising `--max-pages`, or check
  if some pages are genuinely orphaned (no internal links pointing to
  them) and add them via `--extra-urls`.
- **A page's title looks wrong (shows the URL slug instead)** — that page
  timed out or failed to load when the script fetched it; re-run, or check
  the page loads fine in a browser.
