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
MAX_ATTEMPTS = 4             # per page, before giving up on the whole run
RETRY_DELAY = 5
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
    """Fetch one page, retrying both transport errors and holder-less responses.

    The register sometimes answers 200 with a page that has no holder panels at
    all — that is what emptied the 2026-08-01 run, which parsed 0 records in 18s
    and then scraped fine on a re-run. Paging past the end re-serves earlier
    results rather than an empty page, so "no panels" is always anomalous: it is
    retried, and what actually came back is logged, instead of being read as the
    end of the register. Returns None once the attempts are used up.
    """
    params = {"sf_paged": page} if page > 1 else {}
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            r = session.get(BASE_URL, params=params, timeout=60)
            r.raise_for_status()
        except requests.RequestException as exc:
            print(f"\n  page {page} attempt {attempt}/{MAX_ATTEMPTS}: {exc}")
        else:
            soup = BeautifulSoup(r.text, "html.parser")
            if soup.select(".certificate-holder"):
                return soup
            title = soup.title.get_text(strip=True) if soup.title else ""
            print(f"\n  page {page} attempt {attempt}/{MAX_ATTEMPTS}: no holder panels "
                  f"(HTTP {r.status_code}, {len(r.text)} bytes, title {title!r})")
        if attempt < MAX_ATTEMPTS:
            time.sleep(RETRY_DELAY * attempt)
    return None


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
    if first is None:
        raise SystemExit(f"No certificate holders on page 1 after {MAX_ATTEMPTS} attempts — "
                         "the register was unreachable or its layout changed.")
    advertised = min(max_page(first), MAX_PAGES_CAP)   # display only; not a hard stop
    print(f"  ~{advertised} pages of results.")

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
        if soup is None:
            raise SystemExit(f"Page {page} had no certificate holders after {MAX_ATTEMPTS} "
                             f"attempts ({len(records)} read so far) — refusing to commit a "
                             "truncated register.")
        holders = soup.select(".certificate-holder")
        # The pager is windowed, so the true last page only becomes visible as we
        # walk into it; keep the highest number it has ever shown as the check.
        advertised = max(advertised, min(max_page(soup), MAX_PAGES_CAP))
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
        print(f"  Page {page}/~{advertised}: total {len(records)}", end="\r", flush=True)
        if new == 0:                                # only already-seen holders → past the end
            break
        time.sleep(REQUEST_DELAY)

    if not records:
        raise SystemExit("Parsed 0 records — page layout may have changed.")
    # Second net under the empty-page check above: the walk should reach the last
    # page the pager ever advertised. One short is normal — the final page often
    # re-serves earlier holders, which is what ends the loop — but a wider gap
    # means pages went missing.
    if page < advertised - 1:
        raise SystemExit(f"Stopped after {page} pages but the pager advertises {advertised} — "
                         f"only {len(records)} holders read; refusing to commit a partial "
                         "register.")
    if page < advertised:
        print(f"  (stopped at page {page} of the {advertised} advertised — the last page "
              "re-served holders already seen.)")
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
