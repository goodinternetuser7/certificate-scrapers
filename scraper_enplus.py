#!/usr/bin/env python3
"""
Scrapes the ENplus® wood-pellet **producer** register:
  https://enplus-pellets.eu/producer/

The page is a WordPress shell; the register is rendered by the site's own
`enplus_category_list` plugin over admin-ajax, so this is a plain-HTTP scraper
(no browser, no credentials):

  POST /wp-admin/admin-ajax.php
       action=handle_category_search_request
       data=<url-encoded serialisation of #certification-list-form>

returns the results <table> as an HTML fragment. The form's hidden `category`
field selects the register (17 = Producer) and the `certificate_status[]` boxes
are 1=Active, 2=Suspended, 3=Terminated — all three are ticked by default, so
the register covers terminated companies too, not just valid ones. Both the
category id and the status ids are read off the live page rather than hard-coded,
because they are plugin term ids that a site rebuild could renumber.

The list gives ENplus ID / producer / status / country / city / website, and each
row carries a `data-company-id` that the site's "company_info" popup resolves via
a second action:

  POST /wp-admin/admin-ajax.php  action=handle_cert_request  id=<company id>

That popup is the only place the **certification body**, quality classes,
certified activities, legal address, certified sites and approved bag designs
appear, so every company is fetched (a few at a time) and merged into its row.

The register publishes no certificate numbers and no validity dates; a
suspension/termination date ("Since:") is the only date available, and only for
companies that are not Active.

Output columns:
  ENplus ID, Producer, Status, Status Since, Country, City, Certification Body,
  Quality Classes, Certified Activities, Legal Address, Certified Sites,
  Certified Site Names, Bag Designs, Website, Company ID
"""

import csv
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from urllib.parse import urlencode

import requests
from bs4 import BeautifulSoup, NavigableString, Tag

BASE = "https://enplus-pellets.eu"
LIST_URL = f"{BASE}/producer/"
AJAX_URL = f"{BASE}/wp-admin/admin-ajax.php"
LIST_ACTION = "handle_category_search_request"
DETAIL_ACTION = "handle_cert_request"
USER_AGENT = "Mozilla/5.0 ENplus-producer-scraper/1.0"

CATEGORY_FALLBACK = "17"                    # Producer
STATUS_FALLBACK = ("1", "2", "3")           # Active, Suspended, Terminated
PAGE_SIZE = 100                             # the largest the UI offers
MAX_PAGES = int(os.environ.get("ENPLUS_MAX_PAGES", "0")) or None   # test knob
WORKERS = int(os.environ.get("ENPLUS_WORKERS", "6"))
SKIP_DETAILS = os.environ.get("ENPLUS_SKIP_DETAILS") == "1"
MAX_ATTEMPTS = 3
RETRY_DELAY = 3
# A handful of popups may legitimately fail on a flaky day; a large share failing
# means the endpoint changed and the run should not be committed as a scrape.
MAX_DETAIL_FAILURE_RATE = 0.02

# The popup marks each status with an icon class rather than words.
ICON_STATUS = {"icon-ok": "Active", "icon-pause": "Suspended",
               "icon-cancel": "Terminated"}

FIELDNAMES = [
    "ENplus ID", "Producer", "Status", "Status Since", "Country", "City",
    "Certification Body", "Quality Classes", "Certified Activities",
    "Legal Address", "Certified Sites", "Certified Site Names", "Bag Designs",
    "Website", "Company ID",
]

_local = threading.local()


def session():
    """One requests.Session per thread (a Session is not thread-safe)."""
    s = getattr(_local, "session", None)
    if s is None:
        s = requests.Session()
        s.headers["User-Agent"] = USER_AGENT
        s.headers["X-Requested-With"] = "XMLHttpRequest"
        _local.session = s
    return s


def post(data, timeout=120):
    """POST one admin-ajax action, retrying a few times on transport errors."""
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            r = session().post(AJAX_URL, data=data, timeout=timeout)
            r.raise_for_status()
            return r.text
        except requests.RequestException as exc:
            if attempt == MAX_ATTEMPTS:
                raise
            print(f"    retry {attempt}/{MAX_ATTEMPTS - 1} after {exc}")
            time.sleep(RETRY_DELAY * attempt)


# ── The search form ──────────────────────────────────────────────────────────
def discover_form():
    """Read the register's own form ids off the live page.

    Falls back to the known ids if the page is unreachable or restyled — a
    stale category id would silently scrape the wrong register, so the values
    actually used are printed.
    """
    category, statuses = CATEGORY_FALLBACK, list(STATUS_FALLBACK)
    try:
        r = session().get(LIST_URL, timeout=60)
        r.raise_for_status()
        form = BeautifulSoup(r.text, "html.parser").find(id="certification-list-form")
        if form:
            cat = form.find("input", attrs={"name": "category"})
            if cat and cat.get("value"):
                category = cat["value"].strip()
            found = [i["value"] for i in form.find_all("input", attrs={"name": "certificate_status[]"})
                     if i.get("value")]
            if found:
                statuses = found
    except (requests.RequestException, ValueError) as exc:
        print(f"  could not read the form off {LIST_URL} ({exc}); using known ids.")
    return category, statuses


def list_page(category, statuses, page):
    """Fetch one page of the results table (returned as an HTML fragment)."""
    form = [("certificate_status[]", s) for s in statuses] + [
        ("country", ""), ("keyword", ""),
        ("page_size", str(PAGE_SIZE)), ("page_number", str(page)),
        ("sort_order", ""), ("sort_direction", ""), ("category", category),
    ]
    return post({"data": urlencode(form), "action": LIST_ACTION})


def parse_list(html):
    """Rows of the results table → dicts, plus the highest page number offered."""
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for tr in soup.select("tbody tr"):
        cells = tr.find_all("td")
        if len(cells) < 6:
            continue
        link = cells[1].find("a", class_="company_info")
        if not link or not link.get("data-company-id"):
            continue
        site = cells[5].find("a")
        rows.append({
            "ENplus ID": cells[0].get_text(" ", strip=True),
            "Producer": link.get_text(" ", strip=True),
            "Status": cells[2].get_text(" ", strip=True),
            "Country": cells[3].get_text(" ", strip=True),
            "City": cells[4].get_text(" ", strip=True),
            "Website": (site.get("href") or "").strip() if site else "",
            "Company ID": link["data-company-id"].strip(),
        })
    pages = [int(a["data-value"]) for a in soup.select(".pagination a.page-number")
             if (a.get("data-value") or "").isdigit()]
    return rows, max(pages) if pages else 1


def collect_list(category, statuses):
    """Page through the register, unioning rows by company id."""
    html = list_page(category, statuses, 1)
    rows, last_page = parse_list(html)
    if not rows:
        raise SystemExit("Parsed 0 rows from page 1 — the register layout or the "
                         "admin-ajax action may have changed.")
    if MAX_PAGES:
        last_page = min(last_page, MAX_PAGES)
    print(f"  page 1/{last_page}: {len(rows)} rows")

    companies = {r["Company ID"]: r for r in rows}
    for page in range(2, last_page + 1):
        page_rows, _ = parse_list(list_page(category, statuses, page))
        for r in page_rows:
            companies[r["Company ID"]] = r
        print(f"  page {page}/{last_page}: {len(page_rows)} rows, {len(companies)} unique")
        if not page_rows:
            print(f"  WARNING: page {page} came back empty.")

    # The register reports pages, not a total, so the page count is the check:
    # a full last page means there is probably another page we were not offered.
    lo, hi = (last_page - 1) * PAGE_SIZE, last_page * PAGE_SIZE
    if not MAX_PAGES and not lo < len(companies) <= hi:
        print(f"  WARNING: {len(companies)} companies is outside the {lo + 1}–{hi} "
              f"implied by {last_page} pages of {PAGE_SIZE}.")
    return companies


# ── The per-company popup ────────────────────────────────────────────────────
def labelled_blocks(soup):
    """Map each <h6> label → the sibling nodes that follow it, up to the next one.

    The popup is a flat run of <h6>label</h6> + value markup rather than nested
    containers, so a value is "everything between this heading and the next".
    """
    blocks = {}
    for h6 in soup.find_all("h6"):
        key = re.sub(r"[^a-z0-9]+", "_", h6.get_text(" ", strip=True).lower()).strip("_")
        block = []
        for sib in h6.next_siblings:
            if isinstance(sib, Tag) and sib.name == "h6":
                break
            block.append(sib)
        blocks[key] = block
    return blocks


def block_text(block, sep=" "):
    parts = []
    for node in block:
        text = node.strip() if isinstance(node, NavigableString) else node.get_text(sep, strip=True)
        if text:
            parts.append(text)
    return sep.join(parts).strip()


def block_items(block):
    """The <li> texts of any list in the block (quality classes, activities)."""
    items = []
    for node in block:
        if isinstance(node, Tag):
            items += [li.get_text(" ", strip=True) for li in node.find_all("li")]
    return [i for i in items if i]


def block_lines(block):
    """The block's text split on <br> — used for the multi-line legal address."""
    lines = []
    for node in block:
        if isinstance(node, NavigableString):
            if node.strip():
                lines.append(node.strip())
        elif node.name == "br":
            continue
        else:
            lines += [ln.strip() for ln in node.get_text("\n").split("\n") if ln.strip()]
    return lines


def parse_sites(block):
    """Certified sites: '<span class=icon-*></span> Name<br>' pairs.

    A site can carry a different status from its company (e.g. one suspended
    plant of an active producer), so the icon's meaning is kept with the name.
    """
    sites, pending = [], ""
    for node in block:
        if isinstance(node, Tag) and node.name == "span":
            pending = next((ICON_STATUS[c] for c in node.get("class", []) if c in ICON_STATUS), "")
        elif isinstance(node, NavigableString) and node.strip():
            sites.append(f"{node.strip()} ({pending})" if pending else node.strip())
            pending = ""
    return sites


def parse_detail(html):
    soup = BeautifulSoup(html, "html.parser")
    blocks = labelled_blocks(soup)
    if "enplus_company_id" not in blocks:
        # Without its headings the popup can still be parsed into a row of empty
        # strings, which would look like a company with no published detail
        # rather than like a broken endpoint. Fail instead, and let the
        # failure-rate check decide whether it was a blip or a redesign.
        raise ValueError("popup has no 'ENplus® Company ID' heading")

    status_text = block_text(blocks.get("status_of_certificate", []))
    since = re.search(r"Since:\s*(\d{2})-(\d{2})-(\d{4})", status_text)
    sites = parse_sites(blocks.get("certified_sites", []))
    bags = [a["data-elementor-lightbox-title"] for node in blocks.get("approved_bag_designs", [])
            if isinstance(node, Tag)
            for a in node.find_all("a", attrs={"data-elementor-lightbox-title": True})]

    return {
        "ENplus ID": block_text(blocks.get("enplus_company_id", [])),
        "Status Since": f"{since[3]}-{since[2]}-{since[1]}" if since else "",
        "Certification Body": block_text(blocks.get("certification_body", [])),
        "Quality Classes": "; ".join(block_items(blocks.get("quality_classes", []))),
        "Certified Activities": "; ".join(block_items(blocks.get("certified_activities", []))),
        "Legal Address": ", ".join(block_lines(blocks.get("legal_address", []))),
        "Certified Sites": len(sites),
        "Certified Site Names": "; ".join(sites),
        "Bag Designs": len(bags),
    }


def fetch_detail(company_id):
    return parse_detail(post({"id": company_id, "action": DETAIL_ACTION}, timeout=60))


def merge_details(companies):
    """Fetch every company's popup and merge it into that company's row."""
    failed, mismatched, done = [], [], 0
    lock = threading.Lock()

    def work(item):
        nonlocal done
        company_id, row = item
        try:
            detail = fetch_detail(company_id)
        except Exception as exc:            # one odd popup must not kill the run
            with lock:
                failed.append((company_id, f"{type(exc).__name__}: {exc}"[:120]))
            return
        # The popup repeats the ENplus ID, so it doubles as a check that the
        # row's company id really addresses this company.
        popup_id = detail.pop("ENplus ID", "")
        if popup_id and row["ENplus ID"] and popup_id != row["ENplus ID"]:
            with lock:
                mismatched.append((company_id, row["ENplus ID"], popup_id))
        row["ENplus ID"] = row["ENplus ID"] or popup_id
        row.update(detail)
        with lock:
            done += 1
            if done % 200 == 0:
                print(f"    {done}/{len(companies)} popups fetched")

    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        list(pool.map(work, list(companies.items())))

    if failed:
        print(f"  {len(failed)} popup(s) failed, e.g. {failed[:3]}")
        if len(failed) > max(5, MAX_DETAIL_FAILURE_RATE * len(companies)):
            raise SystemExit(f"{len(failed)} of {len(companies)} company popups failed — "
                             "the detail endpoint may have changed.")
    if mismatched:
        raise SystemExit(f"{len(mismatched)} popup(s) reported a different ENplus ID "
                         f"than the list, e.g. {mismatched[:3]} — the company id → row "
                         "mapping is wrong, so the merge cannot be trusted.")
    return failed


def main():
    category, statuses = discover_form()
    print(f"Reading the ENplus producer register (category {category}, "
          f"statuses {','.join(statuses)}) …")
    companies = collect_list(category, statuses)

    if SKIP_DETAILS:
        print("  ENPLUS_SKIP_DETAILS=1 — leaving popup fields empty.")
    else:
        print(f"Fetching {len(companies)} company popups ({WORKERS} at a time) …")
        merge_details(companies)

    records = [{f: row.get(f, "") for f in FIELDNAMES} for row in companies.values()]
    records.sort(key=lambda r: (r["ENplus ID"], r["Producer"].lower()))

    by_status = {}
    for r in records:
        by_status[r["Status"]] = by_status.get(r["Status"], 0) + 1
    bodies = {r["Certification Body"] for r in records if r["Certification Body"]}
    print(f"\nParsed {len(records)} producers ("
          + ", ".join(f"{n} {s or 'Unknown'}" for s, n in sorted(by_status.items()))
          + f") from {len(bodies)} certification bodies.")

    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"ENplus certificates {date_str}.csv", "ENplus certificates latest.csv"):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDNAMES)
            w.writeheader()
            w.writerows(records)
        print(f"Saved → {path}")


if __name__ == "__main__":
    main()
