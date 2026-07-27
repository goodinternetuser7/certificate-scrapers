#!/usr/bin/env python3
"""
Scrapes certified organizations from the FSSC public register:
  https://www.fssc.com/public-register/

The register page is a Vue app on a WordPress site: the table is rendered from a
guest-accessible **admin-ajax** action (`certificate_getCertificates`), so this is
a plain-HTTP scraper (no browser). The action returns JSON with the running
`total` and takes `offset` / `limit`.

Three things shape the pacing — all of them measured against the live site:
  • Cloudflare answers **403** to a non-browser User-Agent, so a browser UA plus
    the register Referer are required.
  • `limit` is capped **server-side at 15** (asking for more still returns 15),
    so the ~42k organizations need ~2.8k requests.
  • Cloudflare rate-limits the endpoint at roughly 75 requests/minute per IP
    (a 429 with `Retry-After: ~125`). Requests are therefore paced one per
    REQUEST_DELAY seconds (1.1s ≈ 55/min, measured as sustainable) — a full run
    takes ~55 minutes. A 429 is not fatal: the scraper honours `Retry-After`,
    then permanently slows itself down before retrying that offset.

The register is offset-paged over a live, alphabetically sorted database that is
updated daily, so a row inserted or removed mid-run shifts every later offset by
one and can duplicate or skip a record at the seam. Records are therefore
de-duplicated by COID and the final count is checked against the API's own
`total`; a shortfall beyond a small tolerance fails the run.

Output columns:
  COID, Organization, Country, City, Scheme, Status, Food Chain Category,
  Product Types, Scope Statement, Initial Certification, Issued, Valid Until,
  Last Status Decision, GFSI Recognized
"""

import csv
import os
import sys
import time
from datetime import datetime, timezone

import requests

AJAX_URL = "https://www.fssc.com/wp-admin/admin-ajax.php"
ACTION = "certificate_getCertificates"
PAGE_SIZE = int(os.environ.get("FSSC_PAGE_SIZE", "15"))     # server caps at 15
REQUEST_DELAY = float(os.environ.get("FSSC_DELAY", "1.1"))  # ≈55 req/min
MAX_PAGES = int(os.environ.get("FSSC_MAX_PAGES", "0"))      # 0 = no cap (local tests)
RETRIES = 5
RETRY_BACKOFF = 5.0          # for network/5xx errors; 429 uses Retry-After
RATE_LIMIT_WAIT = 130.0      # fallback when a 429 carries no Retry-After
RATE_LIMIT_SLOWDOWN = 1.25   # each 429 permanently stretches the delay
# Cloudflare 403s a non-browser UA on this endpoint.
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://www.fssc.com/public-register/",
}
# Records missing at the end of a run, as a fraction of the API's own total,
# that we still accept as live-register drift rather than a broken scrape.
SHORTFALL_TOLERANCE = 0.005

FIELDNAMES = [
    "COID", "Organization", "Country", "City", "Scheme", "Status",
    "Food Chain Category", "Product Types", "Scope Statement",
    "Initial Certification", "Issued", "Valid Until", "Last Status Decision",
    "GFSI Recognized",
]

session = requests.Session()
session.headers.update(HEADERS)
delay = REQUEST_DELAY        # grows if the site rate-limits us
_last_request = 0.0


def throttle():
    """Keep at least `delay` seconds between requests."""
    global _last_request
    wait = delay - (time.monotonic() - _last_request)
    if wait > 0:
        time.sleep(wait)
    _last_request = time.monotonic()


def fetch_page(offset, limit=PAGE_SIZE):
    """Returns (certificates, total) for one offset, retrying transient errors."""
    global delay
    params = {"action": ACTION, "offset": offset, "limit": limit}
    for attempt in range(RETRIES):
        try:
            throttle()
            r = session.get(AJAX_URL, params=params, timeout=60)
            if r.status_code == 429:                    # Cloudflare rate limit
                try:
                    pause = float(r.headers.get("Retry-After", RATE_LIMIT_WAIT))
                except ValueError:
                    pause = RATE_LIMIT_WAIT
                delay *= RATE_LIMIT_SLOWDOWN            # stay slower from here on
                print(f"\n  Rate-limited at offset {offset}; waiting {pause:.0f}s, "
                      f"delay now {delay:.2f}s.", flush=True)
                time.sleep(pause)
                continue
            r.raise_for_status()
            data = r.json()
            return data.get("certificates") or [], int(data.get("total") or 0)
        except (requests.RequestException, ValueError) as e:
            if attempt == RETRIES - 1:
                raise RuntimeError(f"offset {offset} failed after {RETRIES} tries: {e}")
            time.sleep(RETRY_BACKOFF * (attempt + 1))
    raise RuntimeError(f"offset {offset} still rate-limited after {RETRIES} tries")


def category_names(block):
    """`categoryX` blocks are {"categories": [{"name": ...}, ...]} or null."""
    if not block:
        return ""
    return "; ".join(c.get("name", "") for c in block.get("categories") or [] if c.get("name"))


def parse(cert):
    org = cert.get("organisation") or {}
    addr = org.get("address") or {}
    return {
        "COID": cert.get("coid", ""),
        "Organization": org.get("name") or cert.get("title", ""),
        "Country": addr.get("country", ""),
        "City": addr.get("city", ""),
        "Scheme": cert.get("scheme", ""),
        "Status": cert.get("status", ""),
        "Food Chain Category": category_names(cert.get("categoryFoodChain")),
        "Product Types": category_names(cert.get("categoryProductType")),
        # Scope statements carry hard line breaks; flatten so the CSV stays one row.
        "Scope Statement": " ".join((cert.get("scopeStatement") or "").split()),
        "Initial Certification": cert.get("initialCertification", ""),
        "Issued": cert.get("issued", ""),
        "Valid Until": cert.get("validUntil", ""),
        "Last Status Decision": cert.get("lastStatusDecision", ""),
        "GFSI Recognized": cert.get("gfsiRecognized", ""),
    }


def main():
    print(f"Fetching {AJAX_URL}?action={ACTION} …")
    first, total = fetch_page(0)
    if not first or not total:
        raise SystemExit("Register returned no records — the endpoint may have changed.")
    offsets = list(range(PAGE_SIZE, total, PAGE_SIZE))
    if MAX_PAGES:
        offsets = offsets[:max(MAX_PAGES - 1, 0)]
    pages = len(offsets) + 1
    print(f"  {total} certified organizations, {pages} pages of {PAGE_SIZE} "
          f"at {delay:.2f}s/request (~{pages * delay / 60:.0f} min).")

    records, seen, failed = [], set(), []

    def collect(certs, page):
        for c in certs:
            key = c.get("coid") or c.get("id")
            if key in seen:                 # page-seam overlap from live drift
                continue
            seen.add(key)
            records.append(parse(c))
        print(f"  Page {page}/{pages}: {len(records)} records", end="\r", flush=True)

    def try_page(offset):
        """A page that still fails after its retries is parked, not fatal — one
        bad offset shouldn't throw away an hour of good pages."""
        try:
            return fetch_page(offset)[0]
        except RuntimeError as e:
            failed.append((offset, e))
            return []

    collect(first, 1)
    for i, offset in enumerate(offsets, start=2):
        collect(try_page(offset), i)

    if failed:                              # one more pass over the stragglers
        retry, failed = list(failed), []
        print(f"\n  Retrying {len(retry)} failed page(s) …")
        for offset, _ in retry:
            collect(try_page(offset), pages)
        for offset, err in failed:
            print(f"\n  Page at offset {offset} gave up: {err}")

    expected = min(total, pages * PAGE_SIZE)
    missing = expected - len(records)
    print(f"\nParsed {len(records)} certified organizations "
          f"(API total {total}, {missing} short).")
    if missing > max(10, expected * SHORTFALL_TOLERANCE):
        sys.exit(f"Only {len(records)} of {expected} records collected — "
                 "the register or its API likely changed.")

    by_status = {}
    for r in records:
        by_status[r["Status"]] = by_status.get(r["Status"], 0) + 1
    print("  " + ", ".join(f"{k or 'Unknown'}: {v}" for k, v in sorted(by_status.items())))

    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"FSSC certificates {date_str}.csv", "FSSC certificates latest.csv"):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDNAMES)
            w.writeheader()
            w.writerows(records)
        print(f"Saved → {path}")


if __name__ == "__main__":
    main()
