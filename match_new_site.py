#!/usr/bin/env python3
"""
match_new_site.py

Second stage of the redirect-map workflow (run AFTER sitemap_to_redirect_map.py
has already produced the workbook from the OLD site).

  - crawls the NEW site's sitemap (or falls back to link-crawling)
  - for every new-site page, gets its title and its slug (URL path, no domain)
  - if the title exactly matches (case/whitespace-insensitive) a title already
    in Column A, fills that row's Column C with the new-site slug
  - if there's no match, appends a new row at the bottom: title in A,
    Column B left blank, slug in C
  - updates A1 to the new site's URL

Usage:
    python match_new_site.py redirect_map.xlsx https://www.newsite.com/

Requires: requests, beautifulsoup4, openpyxl
    (same as sitemap_to_redirect_map.py, which must be in the same folder)
"""

import argparse
import sys
from urllib.parse import urlparse

from openpyxl import load_workbook

from sitemap_to_redirect_map import (
    find_sitemap_url,
    collect_urls,
    crawl_site,
    is_excluded_url,
    is_homepage_url,
    clean_title,
    fetch_page_data,
    resolve_duplicate_titles,
)


def get_new_site_pages(site_url: str, max_pages: int = 300, no_crawl: bool = False) -> list:
    """Returns a list of (title, slug, url) for every page on the new site."""
    print(f"Locating sitemap for {site_url} ...", file=sys.stderr)
    try:
        sitemap_url = find_sitemap_url(site_url)
        print(f"Using sitemap: {sitemap_url}", file=sys.stderr)
        all_urls = collect_urls(sitemap_url)  # [(url, sitemap_type), ...]
    except RuntimeError as e:
        if no_crawl:
            raise
        print(f"{e}", file=sys.stderr)
        print("Falling back to crawling the site for pages instead...", file=sys.stderr)
        crawled = crawl_site(site_url, max_pages=max_pages)
        all_urls = [(u, None) for u in crawled]
        print(f"Crawled {len(all_urls)} page(s).", file=sys.stderr)

    urls = [(u, t) for u, t in all_urls if not is_excluded_url(u)]

    parsed = urlparse(site_url)
    site_root = f"{parsed.scheme}://{parsed.netloc}/"

    entries = []
    for i, (url, sitemap_type) in enumerate(urls, 1):
        if is_homepage_url(url, site_root):
            title, heading = "Home", None
        elif sitemap_type == "post":
            print(f"  [{i}/{len(urls)}] (blog post per sitemap, skipped) -> {url}", file=sys.stderr)
            continue
        else:
            raw_title, is_post, heading = fetch_page_data(url)
            if sitemap_type is None and is_post:
                print(f"  [{i}/{len(urls)}] (blog post, skipped) -> {url}", file=sys.stderr)
                continue
            title = clean_title(raw_title)
        entries.append({"title": title, "url": url, "heading": heading})
        print(f"  [{i}/{len(urls)}] {title} -> {url}", file=sys.stderr)

    resolve_duplicate_titles(entries)

    pages = []
    for e in entries:
        slug = urlparse(e["url"]).path.strip("/") or "home"
        pages.append((e["title"], slug, e["url"]))

    return pages


def merge_into_workbook(xlsx_path: str, new_site_url: str, pages: list, out_path: str):
    wb = load_workbook(xlsx_path)
    ws = wb.active

    # A1: swap the old site's URL for the new site's, keep the same hyperlink styling
    parsed = urlparse(new_site_url)
    new_root = f"{parsed.scheme}://{parsed.netloc}/"
    ws["A1"] = new_root
    ws["A1"].hyperlink = new_root

    # Build a lookup of existing titles (row 3 downward) -> row number
    title_to_row = {}
    last_row = 2
    row = 3
    while ws.cell(row=row, column=1).value:
        title = str(ws.cell(row=row, column=1).value).strip()
        title_to_row[title.lower()] = row
        last_row = row
        row += 1

    matched, appended = 0, 0
    next_row = last_row + 1

    for title, slug, _url in pages:
        key = title.strip().lower()
        if key in title_to_row:
            ws.cell(row=title_to_row[key], column=3, value=slug)
            matched += 1
        else:
            ws.cell(row=next_row, column=1, value=title)
            ws.cell(row=next_row, column=3, value=slug)
            title_to_row[key] = next_row
            next_row += 1
            appended += 1

    wb.save(out_path)
    print(f"Matched {matched} page(s) into existing rows, appended {appended} new row(s).", file=sys.stderr)
    print(f"Wrote {out_path}", file=sys.stderr)


def main():
    try:
        import lxml  # noqa: F401
    except ImportError:
        sys.exit(
            "Missing dependency: 'lxml' is required to parse sitemap.xml files.\n"
            "Install it with:  pip install lxml"
        )

    ap = argparse.ArgumentParser(description="Match a new site's pages into an existing redirect-map workbook.")
    ap.add_argument("workbook", help="Path to the .xlsx produced by sitemap_to_redirect_map.py")
    ap.add_argument("new_site", help="New site's base URL or sitemap.xml URL")
    ap.add_argument("-o", "--output", default=None,
                     help="Output .xlsx path (default: overwrite the input workbook)")
    ap.add_argument("--max-pages", type=int, default=300,
                     help="Max pages to visit when crawling a site with no sitemap (default: 300)")
    ap.add_argument("--no-crawl", action="store_true",
                     help="Don't fall back to crawling the site if no sitemap is found")
    args = ap.parse_args()

    out_path = args.output or args.workbook

    pages = get_new_site_pages(args.new_site, max_pages=args.max_pages, no_crawl=args.no_crawl)
    merge_into_workbook(args.workbook, args.new_site, pages, out_path)


if __name__ == "__main__":
    main()