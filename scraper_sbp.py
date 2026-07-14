#!/usr/bin/env python3
"""
Scrapes certificate holders from the Sustainable Biomass Program (SBP) register:
  https://sbp-cert.org/certificate-holders/

The register is a WordPress "Search & Filter Pro" directory: each holder is a
server-rendered expandable panel, 12 per page, paginated via the `?sf_paged=N`
query parameter. This is a plain-HTTP scraper (no browser) that reads the max
page number from the first page's pagination, then walks every page and parses
each panel's detail block (a clean <span class="label">/<span class="value">
list) plus the holder name and country flag from the header.

Output columns:
  Certificate Number, Certificate Holder, Country, Certificate Type, Status,
  Certification Body, Date of Issue, Date of Expiry, Certificate Scope,
  Products Covered
"""

import csv
import re
import time
from datetime import datetime, timezone

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://sbp-cert.org/certificate-holders/"
REQUEST_DELAY = 0.3
MAX_PAGES_CAP = 500          # backstop against a runaway paging loop
USER_AGENT = "Mozilla/5.0 SBP-cert-scraper/1.0"

FIELDNAMES = [
    "Certificate Number", "Certificate Holder", "Country", "Certificate Type",
    "Status", "Certification Body", "Date of Issue", "Date of Expiry",
    "Certificate Scope", "Products Covered",
]
# Detail-block labels we keep, mapped to their output column name.
DETAIL_FIELDS = {
    "Certification Body": "Certification Body",
    "Certificate Number": "Certificate Number",
    "Status": "Status",
    "Date of Issue": "Date of Issue",
    "Date of Expiry": "Date of Expiry",
    "Certificate Type": "Certificate Type",
    "Certificate Scope": "Certificate Scope",
    "Products Covered": "Products Covered",
}


def get_soup(session, page):
    params = {"sf_paged": page} if page > 1 else {}
    r = session.get(BASE_URL, params=params, timeout=60)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser")


def max_page(soup):
    pages = [int(m) for m in re.findall(r"sf_paged=(\d+)", str(soup))]
    return max(pages) if pages else 1


def parse_holder(h):
    rec = {k: "" for k in FIELDNAMES}

    # Name: the header anchor text with the certificate-number span removed.
    name_el = h.select_one(".certificate-holder-name")
    if name_el:
        num = name_el.select_one(".certificate-holder-number")
        num_txt = num.get_text(strip=True) if num else ""
        full = name_el.get_text(" ", strip=True)
        rec["Certificate Holder"] = full.replace(num_txt, "", 1).strip()

    # Country: from the header flag image alt ("Flag of Switzerland").
    flag = h.select_one("img.country-flag")
    if flag and flag.get("alt"):
        rec["Country"] = re.sub(r"^Flag of\s+", "", flag["alt"]).strip()

    # Detail fields: <li><span class="label">X: </span><span class="value">Y</span></li>
    ci = h.select_one(".certificate-information ul.certification-list-meta")
    if ci:
        for li in ci.find_all("li", recursive=False):
            label = li.select_one("span.label")
            if not label:
                continue
            key = label.get_text(strip=True).rstrip(":").strip()
            if key not in DETAIL_FIELDS:
                continue
            val_ul = li.find("ul", class_="value")
            if val_ul:                              # multi-value (Products / Scope)
                val = "; ".join(x.get_text(strip=True) for x in val_ul.find_all("li"))
            else:
                v = li.select_one("span.value")
                val = v.get_text(strip=True) if v else ""
            rec[DETAIL_FIELDS[key]] = val
    return rec


def main():
    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT

    print(f"Fetching {BASE_URL} …")
    first = get_soup(session, 1)
    shown = min(max_page(first), MAX_PAGES_CAP)     # display only; not a hard stop
    print(f"  ~{shown} pages of results.")

    # Walk pages until one yields no new holders. We don't trust max_page as the
    # loop bound: Search & Filter Pro renders a *windowed* pager (1 2 3 … Next)
    # that can omit the true last-page link, which would silently truncate. When
    # sf_paged runs past the end SF Pro re-serves earlier results, so every
    # holder is already seen and "zero new" ends the loop cleanly.
    records, seen = [], set()
    page = 0
    while page < MAX_PAGES_CAP:
        page += 1
        soup = first if page == 1 else get_soup(session, page)
        holders = soup.select(".certificate-holder")
        if not holders:
            break
        new = 0
        for h in holders:
            rec = parse_holder(h)
            key = rec["Certificate Number"] or (
                rec["Certificate Holder"], rec["Certification Body"], rec["Date of Issue"])
            if key in seen:                         # guard against page overlap / overrun
                continue
            seen.add(key)
            records.append(rec)
            new += 1
        print(f"  Page {page}/~{shown}: total {len(records)}", end="\r", flush=True)
        if new == 0:                                # only already-seen holders → past the end
            break
        time.sleep(REQUEST_DELAY)

    if not records:
        raise SystemExit("Parsed 0 records — page layout may have changed.")
    active = sum(1 for r in records if r["Status"].lower() == "active")
    print(f"\nParsed {len(records)} certificate holders ({active} Active).")

    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"SBP certificates {date_str}.csv", "SBP certificates latest.csv"):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDNAMES)
            w.writeheader()
            w.writerows(records)
        print(f"Saved → {path}")


if __name__ == "__main__":
    main()
