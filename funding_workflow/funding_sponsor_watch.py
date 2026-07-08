#!/usr/bin/env python3
"""
funding_sponsor_watch.py

Weekly pipeline for the "hiring gap" job-search strategy:

  1. Scrapes UKTN's funding news feed (https://www.uktech.news/funding) for
     recently-funded UK tech companies.
  2. Downloads the current Home Office "Register of licensed sponsors: workers"
     CSV and cross-references each funded company against it.
  3. Reports:
       - Companies that are ALREADY a licensed sponsor (best cold-outreach
         targets: funded + can legally hire you right now)
       - Companies that are NOT on the register yet (still worth a warm
         "advice call" reach-out, but don't expect sponsorship on day one)
  4. Keeps a local cache (seen_articles.json) so re-runs only surface NEW
     funding stories since the last run -- run this weekly (e.g. via cron
     or a calendar reminder) and it behaves like a standing monitor.

Usage:
    python funding_sponsor_watch.py                  # default: scan 3 pages
    python funding_sponsor_watch.py --pages 5
    python funding_sponsor_watch.py --reset-cache     # forget previously seen articles
    python funding_sponsor_watch.py --min-rounds 20   # keep scanning pages until N rounds found

Requires: requests  (pip install requests --break-system-packages)
"""

import argparse
import csv
import io
import json
import re
import sys
import unicodedata
from pathlib import Path

import requests
from bs4 import BeautifulSoup

UKTN_FUNDING_URL = "https://www.uktech.news/funding"
UKTN_FUNDING_PAGE_URL = "https://www.uktech.news/funding/page/{n}"
GOVUK_REGISTER_PAGE = "https://www.gov.uk/government/publications/register-of-licensed-sponsors-workers"

CACHE_FILE = Path(__file__).parent / "seen_articles.json"
OUTPUT_CSV = Path(__file__).parent / "funding_sponsor_matches.csv"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; JobSearchResearchBot/1.0; personal use)"
}

# Verbs UKTN headlines use for a raise, so we can split "Company <verb> £Xm ..."
RAISE_VERBS = [
    "raises", "raise", "secures", "score", "scores", "closes", "close",
    "lands", "bags", "nets", "wins", "completes", "banks", "pulls in",
    "receives", "gets", "confirms",
]

VERB_PATTERN = re.compile(
    r"^(?P<company>.+?)\s+(?:" + "|".join(RAISE_VERBS) + r")\s+£", re.IGNORECASE
)

# Words to strip when normalising a company name for matching purposes
LEGAL_SUFFIXES = re.compile(
    r"\b(limited|ltd|llp|plc|inc|incorporated|group|holdings|technologies|"
    r"technology|labs|lab|uk|co|company)\b\.?",
    re.IGNORECASE,
)
NON_ALNUM = re.compile(r"[^a-z0-9 ]")


# --------------------------------------------------------------------------
# Step 1: scrape UKTN funding listing pages for article titles
# --------------------------------------------------------------------------

def fetch_uktn_page(page_num: int) -> str:
    url = UKTN_FUNDING_URL if page_num == 1 else UKTN_FUNDING_PAGE_URL.format(n=page_num)
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.text

# UKTN article URLs consistently end in an 8-digit publish date, e.g.
# https://www.uktech.news/fashion-tech/fleek-raises-18-7m-....-20260708
# Matching on that is far more robust than relying on exact heading markup,
# which can change with theme/CMS updates.
ARTICLE_URL_PATTERN = re.compile(r"^https://www\.uktech\.news/[^?#]+-\d{8}/?$")

# Category/nav links and other boilerplate we don't want to mistake for articles
SKIP_URL_SUBSTRINGS = ("/tech-hubs/", "/category/", "/page/")


def parse_articles(html: str):
    soup = BeautifulSoup(html, "html.parser")
    seen_urls = {}
    for a in soup.find_all("a", href=True):
        href = a["href"].split("#")[0]
        if not ARTICLE_URL_PATTERN.match(href):
            continue
        if any(s in href for s in SKIP_URL_SUBSTRINGS):
            continue
        text = a.get_text(strip=True)
        if not text:
            continue
        # Same article often appears twice (image link + title link);
        # keep whichever anchor text is longest/most title-like.
        if href not in seen_urls or len(text) > len(seen_urls[href]):
            seen_urls[href] = text

    return [{"title": re.sub(r"\s+", " ", title).strip(), "url": url}
            for url, title in seen_urls.items()]


def collect_funding_articles(pages: int, min_rounds: int | None):
    all_articles = []
    seen_urls = set()
    page = 1
    while page <= pages or (min_rounds and len(all_articles) < min_rounds):
        try:
            html = fetch_uktn_page(page)
        except requests.RequestException as e:
            print(f"  [warn] could not fetch UKTN page {page}: {e}", file=sys.stderr)
            break
        found = parse_articles(html)
        if not found:
            break
        new = [a for a in found if a["url"] not in seen_urls]
        for a in new:
            seen_urls.add(a["url"])
        all_articles.extend(new)
        if not new:
            # no new articles on this page -> probably hit the end / duplicate page
            break
        page += 1
        if page > 50:  # hard safety cap
            break
        if not min_rounds and page > pages:
            break
    return all_articles


# --------------------------------------------------------------------------
# Step 2: extract a company name from each headline
# --------------------------------------------------------------------------

# UKTN headlines often prefix the real company name with a descriptor, e.g.
# "Cambridge AI robotics group Dogtooth scores..." or
# "London hyperscaler Nscale secures...". Strip everything up to and
# including the last such descriptor word so we're left with just the name.
DESCRIPTOR_STOPWORDS = {
    "startup", "group", "firm", "platform", "company", "maker", "developer",
    "hyperscaler", "robotics", "fintech", "lab", "labs", "venture", "ventures",
    "business", "team", "outfit", "provider", "supplier", "giant", "unicorn",
    "scaleup", "scale-up",
}


def strip_descriptor_prefix(company: str) -> str:
    words = company.split()
    last_stopword_idx = None
    for i, w in enumerate(words):
        if w.lower().strip(".,'") in DESCRIPTOR_STOPWORDS:
            last_stopword_idx = i
    if last_stopword_idx is not None and last_stopword_idx + 1 < len(words):
        return " ".join(words[last_stopword_idx + 1:])
    return company


def extract_company_name(title: str) -> str | None:
    company = None
    m = VERB_PATTERN.search(title)
    if m:
        company = m.group("company").strip()
    else:
        # Fallback: look for a capitalised token cluster right before a raise verb
        m2 = re.search(
            r"([A-Z][A-Za-z0-9&.\-]*(?:\s+[A-Z][A-Za-z0-9&.\-]*){0,3})\s+(?:"
            + "|".join(RAISE_VERBS) + r")\b",
            title,
        )
        if m2:
            company = m2.group(1).strip()

    if not company:
        return None
    return strip_descriptor_prefix(company)


# --------------------------------------------------------------------------
# Step 3: download + index the Home Office sponsor register
# --------------------------------------------------------------------------

def find_register_csv_url() -> str:
    resp = requests.get(GOVUK_REGISTER_PAGE, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    m = re.search(
        r'href="(https://assets\.publishing\.service\.gov\.uk/media/[^"]+\.csv)"',
        resp.text,
    )
    if not m:
        raise RuntimeError(
            "Could not locate the CSV download link on the GOV.UK register page. "
            "The page layout may have changed -- check "
            + GOVUK_REGISTER_PAGE
            + " manually."
        )
    return m.group(1)


def normalise(name: str) -> str:
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    name = name.lower()
    # split off "T/A" or "trading as" trading names -- keep the part before it,
    # since that's usually the legal entity name UKTN / the register both use
    name = re.split(r"\bt/a\b|\btrading as\b", name)[0]
    name = LEGAL_SUFFIXES.sub("", name)
    name = NON_ALNUM.sub("", name)
    name = re.sub(r"\s+", "", name)
    return name.strip()


def download_sponsor_register() -> dict:
    print("Fetching current Home Office sponsor register (this file is ~10MB)...")
    csv_url = find_register_csv_url()
    resp = requests.get(csv_url, headers=HEADERS, timeout=120)
    resp.raise_for_status()
    # Register is usually UTF-8 with a BOM
    text = resp.content.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))

    # Column name can vary slightly release to release; find it defensively
    fieldnames = reader.fieldnames or []
    name_col = next(
        (f for f in fieldnames if "organisation" in f.lower() and "name" in f.lower()),
        None,
    )
    if not name_col:
        raise RuntimeError(f"Could not find an organisation name column. Columns were: {fieldnames}")

    index = {}
    for row in reader:
        org_name = row.get(name_col, "").strip()
        if not org_name:
            continue
        key = normalise(org_name)
        if key and key not in index:
            index[key] = row
    print(f"Indexed {len(index):,} licensed sponsor organisations.")
    return index


# --------------------------------------------------------------------------
# Step 4: cross-reference + cache of already-seen articles
# --------------------------------------------------------------------------

def load_cache() -> set:
    if CACHE_FILE.exists():
        try:
            return set(json.loads(CACHE_FILE.read_text()))
        except Exception:
            return set()
    return set()


def save_cache(urls: set):
    CACHE_FILE.write_text(json.dumps(sorted(urls), indent=2))


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--pages", type=int, default=3, help="How many UKTN funding pages to scan (default 3, ~10 articles/page)")
    parser.add_argument("--min-rounds", type=int, default=None, help="Keep scanning pages until at least this many funding articles are found")
    parser.add_argument("--reset-cache", action="store_true", help="Ignore/clear the 'already seen' cache and report on everything found")
    args = parser.parse_args()

    if args.reset_cache and CACHE_FILE.exists():
        CACHE_FILE.unlink()

    seen = load_cache()

    print(f"Scanning UKTN funding news ({args.pages} page(s))...")
    articles = collect_funding_articles(args.pages, args.min_rounds)
    print(f"Found {len(articles)} funding articles total.")

    new_articles = [a for a in articles if a["url"] not in seen]
    print(f"{len(new_articles)} are new since the last run.\n")

    if not new_articles:
        print("Nothing new this week. Re-run with --reset-cache to see the full current list again.")
        return

    register = download_sponsor_register()

    rows = []
    likely_sponsors = []
    needs_check = []

    for a in new_articles:
        company = extract_company_name(a["title"])
        if not company:
            continue
        key = normalise(company)
        match = register.get(key)

        row = {
            "company": company,
            "headline": a["title"],
            "url": a["url"],
            "sponsor_match": "",
            "sponsor_town": "",
            "sponsor_rating": "",
        }

        if match:
            row["sponsor_match"] = match.get("Organisation Name", "")
            row["sponsor_town"] = match.get("Town/City", "")
            row["sponsor_rating"] = match.get("Type & Rating", "")
            likely_sponsors.append(row)
        else:
            needs_check.append(row)

        rows.append(row)

    # write / append CSV for the user's own records
    write_header = not OUTPUT_CSV.exists()
    with open(OUTPUT_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["company", "headline", "url", "sponsor_match", "sponsor_town", "sponsor_rating"])
        if write_header:
            writer.writeheader()
        writer.writerows(rows)

    print("=" * 70)
    print(f"ALREADY A LICENSED SPONSOR ({len(likely_sponsors)}) -- best targets, funded + can sponsor now")
    print("=" * 70)
    for r in likely_sponsors:
        print(f"  * {r['company']}  [{r['sponsor_town']}, {r['sponsor_rating']}]")
        print(f"      {r['headline']}")
        print(f"      {r['url']}\n")

    print("=" * 70)
    print(f"NOT ON THE REGISTER YET ({len(needs_check)}) -- verify manually / still worth a warm reach-out")
    print("=" * 70)
    for r in needs_check:
        print(f"  * {r['company']}")
        print(f"      {r['headline']}")
        print(f"      {r['url']}\n")

    print(f"Full results appended to: {OUTPUT_CSV}")

    save_cache(seen | {a["url"] for a in new_articles})


if __name__ == "__main__":
    main()
