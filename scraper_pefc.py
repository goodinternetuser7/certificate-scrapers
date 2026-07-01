#!/usr/bin/env python3
"""
Scrapes active (valid) PEFC certificates & licences from the public
"Find certified" search at https://pefc.org/find-certified-legacy

The search is a Caspio DataPage embedded on the PEFC site. It renders only in a
browser (async JS embed), so we drive it with headless Chromium via Playwright:
filter Status = Valid, set 250 results per page, then page through the whole set.

Output columns (what the public list view exposes):
    Code, Entity, City, Country, Role, Certificate Number, Licence Number,
    Category, Status, Type, Entity ID, Certificate ID, Licence ID

Note: the public list does NOT expose the issuing Certification Body or the
expiry date — those live only on per-entity detail pages (one request each,
impractical for ~66k rows), so they are intentionally omitted.

Env vars:
    PEFC_MAX_PAGES   limit pages (for local testing); unset = all pages
    PEFC_PAGE_SIZE   results per page (default 250; Caspio caps at 250)
    PEFC_HEADFUL     any value -> run with a visible browser (debugging)
"""

import csv
import os
import re
import sys
from datetime import datetime, timezone

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

DATAPAGE_URL = "https://c2abx025.caspio.com/dp/30b08000892b061a8b754b248392"
STATUS_FIELD = "Value12_1"          # overviews_id_certificate_statu
STATUS_VALID = "1"
COUNTRY_FIELD = "Value6_1"          # entities_id_country (used to identify the country line)
PAGE_SIZE = os.environ.get("PEFC_PAGE_SIZE", "250")
MAX_PAGES = int(os.environ["PEFC_MAX_PAGES"]) if os.environ.get("PEFC_MAX_PAGES") else None
CHECKPOINT_EVERY = 20               # flush a partial CSV every N pages (crash safety)

FIELDNAMES = [
    "Code", "Entity", "City", "Country", "Role",
    "Certificate Number", "Licence Number",
    "Category", "Status", "Type",
    "Entity ID", "Certificate ID", "Licence ID",
]

ROW_SELECTOR = ("tr.cbResultSetDataRow, tr.cbResultSetTableEvenRow, "
                "tr.cbResultSetTableOddRow")

# Caspio hides its container with an anti-flicker `display:none` that the embed
# host normally removes; loading the DataPage standalone leaves it hidden, so we
# strip it ourselves after every (re)render.
UNHIDE_JS = """() => {
    document.querySelectorAll('[id^=cbOuterAjaxCtnr]').forEach(e => { e.style.display = ''; });
    document.querySelectorAll('style').forEach(s => {
        if (s.textContent && s.textContent.includes('cbOuterAjaxCtnr')) s.remove();
    });
}"""

ROLE_HINTS = ("certificate holder", "licence holder", "license holder", "site/member")

# The Caspio result table is a responsive layout whose cell→column mapping is not
# stable across renders (a status value can land in the "Certificate" cell, the
# number in the "Licence" cell, etc.). So rather than trust cell positions we
# classify each line of the certificate/licence cells by its content. A number line
# always contains a digit; a category ("COC - Multisite", "D - Other") or status
# never does — that distinction handles PEFC and every endorsed national scheme
# (SGEC, CFCC, KFCC, SFI, …) and plain-numeric certificate formats alike.
STATUS_VALUES = {
    "valid", "expired", "suspended", "withdrawn",
    "terminated (voluntary cancellation)", "not pefc recognized",
}
# A logo-licence number looks like SCHEME/digits (PEFC/14-44-00024, SGEC/31-44-1528);
# anything else numeric is a certificate number (SA-PEFC-COC-013288, 8305-09, …).
LICENCE_RE = re.compile(r"^[A-Za-z]{2,}/\d")
HAS_DIGIT = re.compile(r"\d")


# Extract every result row on the current page in ONE browser round-trip. Cells are
# identified by their label (Code / Entity), not by position, because the Caspio
# responsive layout shuffles columns. Everything that isn't Code/Entity/noise is
# pooled into `rest` for content-based classification in build_record().
EXTRACT_JS = """() => {
    const cellLines = (td) => {
        const label = td.querySelector('.cbResultSetLabel');
        let t = (td.innerText || '').trim();
        if (label) {
            const lt = label.innerText.trim();
            if (t.startsWith(lt)) t = t.slice(lt.length);
        }
        return t.split('\\n').map(s => s.trim()).filter(Boolean);
    };
    const rows = document.querySelectorAll(
        'tr.cbResultSetDataRow, tr.cbResultSetTableEvenRow, tr.cbResultSetTableOddRow');
    const out = [];
    for (const r of rows) {
        let code = '', entity = [], rest = [], onclick = '';
        for (const td of r.querySelectorAll('td')) {
            const lab = td.querySelector('.cbResultSetLabel');
            const key = lab ? lab.innerText.trim().replace(/:$/, '') : '';
            const lines = cellLines(td);
            if (key === 'Code') { if (lines.length && !code) code = lines[0]; }
            else if (key === 'Entity') { if (lines.length) entity = lines; }
            else if (key === 'overview id role') { /* internal noise */ }
            else { for (const l of lines) rest.push(l); }
        }
        for (const b of r.querySelectorAll('button')) {
            const oc = b.getAttribute('onclick') || '';
            if (oc.includes('find-certified/details')) { onclick = oc; break; }
        }
        out.push({ code, entity, rest, onclick });
    }
    return out;
}"""


def extract_page(page):
    return page.evaluate(EXTRACT_JS)


def build_record(rec, countries):
    entity = rec["entity"]

    # Entity block lines: name, [address/city…], country, [relationship role].
    # The country is whichever line matches a known PEFC country name; anything
    # after it (e.g. "Certificate holder", "Site/Member") is the role.
    name = entity[0] if entity else ""
    tail = entity[1:]
    country = city = role = ""
    cidx = next((i for i, ln in enumerate(tail) if ln in countries), None)
    if cidx is not None:
        country = tail[cidx]
        city = " ".join(tail[:cidx])
        role = " ".join(tail[cidx + 1:])
    elif tail and tail[-1].lower() in ROLE_HINTS:
        role = tail[-1]
        city = " ".join(tail[:-1])
    else:
        city = " ".join(tail)

    # Classify the remaining cell lines by content (position-independent): numbers
    # carry a digit; status/category do not.
    cert_no = lic_no = category = status = ""
    for ln in rec["rest"]:
        if HAS_DIGIT.search(ln):
            if LICENCE_RE.match(ln):
                lic_no = lic_no or ln
            else:
                cert_no = cert_no or ln
        elif ln.lower() in STATUS_VALUES:
            status = status or ln
        elif " - " in ln:                    # categories look like "COC - Multisite"
            category = category or ln
        # else: stray codes ("AAC", "A", …) — not a field we keep

    def grab(key):
        m = re.search(rf"{key}=([^&'\"]*)", rec["onclick"])
        return m.group(1) if m else ""

    cid, lid, eid = grab("CID"), grab("LID"), grab("EID")
    # The details-link ids are the authoritative signal for what the entity holds
    # (a certificate, a logo licence, or both); fall back to the parsed numbers.
    if cid:
        rec_type = "Certificate"
    elif lid:
        rec_type = "Licence"
    else:
        rec_type = "Certificate" if cert_no else ("Licence" if lic_no else "")

    return {
        "Code": rec["code"],
        "Entity": name,
        "City": city,
        "Country": country,
        "Role": role,
        "Certificate Number": cert_no,
        "Licence Number": lic_no,
        "Category": category,
        "Status": status,
        "Type": rec_type,
        "Entity ID": eid,
        "Certificate ID": cid,
        "Licence ID": lid,
    }


def wait_locker_clear(page):
    """Wait until Caspio's loading overlay (.UILocker) is gone/hidden."""
    try:
        page.wait_for_function(
            """() => {
                const l = document.querySelector('.UILocker');
                return !l || getComputedStyle(l).display === 'none' || l.offsetParent === null;
            }""",
            timeout=60000,
        )
    except PWTimeout:
        pass


def settle(page):
    """Wait for a rendered result set and strip Caspio's hide styles."""
    try:
        page.wait_for_load_state("networkidle", timeout=60000)
    except PWTimeout:
        pass
    page.wait_for_selector(ROW_SELECTOR, state="attached", timeout=60000)
    wait_locker_clear(page)
    page.evaluate(UNHIDE_JS)


def total_count(page):
    m = re.search(r"Records?\s*[\d,]+\s*-\s*[\d,]+\s*of\s*([\d,]+)",
                  page.inner_text("body"))
    return int(m.group(1).replace(",", "")) if m else None


def set_page_size(page, size):
    """The records-per-page <select> has no name; identify it by its exact option
    set (10/25/50/100/250) so we don't hit the country filter, then set it via a
    JS change event (Caspio's loading overlay blocks a native click)."""
    wait_locker_clear(page)
    return page.evaluate(
        """(size) => {
            const SIG = ['10', '25', '50', '100', '250'];
            for (const s of document.querySelectorAll('select')) {
                const vals = [...s.options].map(o => o.value);
                const isPageSize = vals.length && vals.every(v => SIG.includes(v)) && vals.includes('250');
                if (isPageSize && vals.includes(size)) {
                    s.value = size;
                    s.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }
            }
            return false;
        }""",
        size,
    )


# First result row's Code — a stable per-page fingerprint used to confirm that a
# "next" click actually rendered a new page (independent of the flaky counter).
FIRST_CODE_JS = """() => {
    const r = document.querySelector('tr.cbResultSetDataRow, tr.cbResultSetTableEvenRow, tr.cbResultSetTableOddRow');
    if (!r) return '';
    for (const td of r.querySelectorAll('td')) {
        const lab = td.querySelector('.cbResultSetLabel');
        if (lab && lab.innerText.trim().replace(/:$/, '') === 'Code') {
            let t = (td.innerText || '').trim();
            const lt = lab.innerText.trim();
            if (t.startsWith(lt)) t = t.slice(lt.length);
            return t.trim();
        }
    }
    return '';
}"""


def first_code(page):
    return page.evaluate(FIRST_CODE_JS)


def click_next(page):
    wait_locker_clear(page)
    return page.evaluate("""() => {
        for (const a of document.querySelectorAll('a.cbResultSetNavigationLinks')) {
            if (a.innerHTML.includes('set5_next.gif')) { a.click(); return true; }
        }
        return false;
    }""")


def advance(page, prev_code, attempts=4):
    """Click 'next' and wait — using the first row's Code, not the flaky counter —
    until a genuinely new page has rendered. Returns the new first Code on success,
    None if there is no next page / it never advanced. Caspio still occasionally
    skips a page here; the union-of-passes in main() is what makes the run complete."""
    for _ in range(attempts):
        if not click_next(page):
            return None
        try:
            page.wait_for_function(
                """(prev) => {
                    const r = document.querySelector('tr.cbResultSetDataRow, tr.cbResultSetTableEvenRow, tr.cbResultSetTableOddRow');
                    if (!r) return false;
                    let code = '';
                    for (const td of r.querySelectorAll('td')) {
                        const lab = td.querySelector('.cbResultSetLabel');
                        if (lab && lab.innerText.trim().replace(/:$/, '') === 'Code') {
                            let t = (td.innerText || '').trim();
                            const lt = lab.innerText.trim();
                            if (t.startsWith(lt)) t = t.slice(lt.length);
                            code = t.trim();
                            break;
                        }
                    }
                    return code && code !== prev;
                }""",
                arg=prev_code, timeout=45000)
        except PWTimeout:
            pass
        wait_locker_clear(page)
        try:
            page.wait_for_selector(ROW_SELECTOR, state="attached", timeout=30000)
        except PWTimeout:
            pass
        page.evaluate(UNHIDE_JS)
        cur = first_code(page)
        if cur and cur != prev_code:
            return cur
    return None


def run(page):
    print("Loading DataPage …")
    page.goto(DATAPAGE_URL, wait_until="networkidle", timeout=90000)
    page.wait_for_selector("#caspioform", state="attached", timeout=60000)
    page.evaluate(UNHIDE_JS)

    # Known country names, to identify the country line within an entity block.
    countries = set(page.eval_on_selector_all(
        f"select[name={COUNTRY_FIELD}] option",
        "opts => opts.map(o => o.textContent.trim())"))
    countries.discard("- Any -")
    countries.discard("")

    print("Filtering Status = Valid …")
    page.select_option(f"select[name={STATUS_FIELD}]", STATUS_VALID)
    page.eval_on_selector("input[id^=searchID]", "el => el.click()")
    settle(page)

    total = total_count(page)
    print(f"Valid records: {total}")

    page_size = int(PAGE_SIZE)
    if total and total > 25:
        # Setting the page size is flaky — it can report success yet leave the table
        # at 25 rows. Verify it actually took effect and retry, because pageing 50k
        # records at 25/page cannot finish in the workflow window.
        applied = False
        for _ in range(5):
            print(f"Setting results per page = {PAGE_SIZE} …")
            if not set_page_size(page, PAGE_SIZE):
                raise RuntimeError("page-size selector not found")
            settle(page)
            try:
                page.wait_for_function(
                    "(n) => document.querySelectorAll('tr.cbResultSetDataRow, "
                    "tr.cbResultSetTableEvenRow, tr.cbResultSetTableOddRow').length > n",
                    arg=25, timeout=30000)
            except PWTimeout:
                pass
            settle(page)
            # Verify with the same extractor the loop uses (element counts can lie
            # while the table is mid-transition).
            if len(extract_page(page)) > 25:
                applied = True
                break
            print("  page size did not take effect; retrying …", file=sys.stderr)
        if not applied:
            raise RuntimeError("could not set results-per-page to 250")

    records = []
    page_num = 0
    prev_code = first_code(page)
    while True:
        page_num += 1
        page.wait_for_selector(ROW_SELECTOR, state="attached", timeout=60000)
        page_rows = extract_page(page)
        # Partial render guard: a page that is short but not the last one (a "next"
        # arrow still exists) is mid-render — re-settle and re-extract once.
        if len(page_rows) < page_size and page.evaluate(
                "() => [...document.querySelectorAll('a.cbResultSetNavigationLinks')]"
                ".some(a => a.innerHTML.includes('set5_next.gif'))"):
            settle(page)
            page_rows = extract_page(page)

        for rec in page_rows:
            records.append(build_record(rec, countries))
        latest = total_count(page)
        if latest:
            total = latest
        print(f"  Page {page_num}: rows={len(page_rows)} "
              f"collected={len(records)}", end="\r", flush=True)

        if page_num % CHECKPOINT_EVERY == 0:
            write_csv(dedupe(records), announce=False)   # crash-safety checkpoint

        if MAX_PAGES and page_num >= MAX_PAGES:
            print(f"\nReached PEFC_MAX_PAGES={MAX_PAGES}.")
            break
        # The genuine end of the result set is a short (partial) page.
        if len(page_rows) < page_size:
            print(f"\nReached final (partial) page {page_num} "
                  f"with {len(page_rows)} rows.")
            break
        new_code = advance(page, prev_code)
        if new_code is None:
            print(f"\nNo next page after page {page_num}.")
            break
        prev_code = new_code

    records = dedupe(records)
    # The counter is unreliable mid-run but reliable once the final page has fully
    # settled; treat the run as complete only if we actually hold ~all the rows it
    # reports. A page skipped mid-run fails this check and triggers a whole retry.
    final_total = total_count(page) or total
    complete = bool(MAX_PAGES) or bool(
        final_total and len(records) >= final_total - page_size)
    return records, final_total, complete


def dedupe(records):
    """Drop rows with a duplicate Code (guards against a re-extracted page)."""
    seen, out = set(), []
    for r in records:
        code = r["Code"]
        if code and code in seen:
            continue
        seen.add(code)
        out.append(r)
    return out


def write_csv(records, announce=True):
    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"PEFC certificates {date_str}.csv", "PEFC certificates latest.csv"):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDNAMES)
            w.writeheader()
            w.writerows(records)
        if announce:
            print(f"Saved → {path}")


MAX_ATTEMPTS = 6


def scrape_once():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not os.environ.get("PEFC_HEADFUL"))
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        try:
            return run(page)
        finally:
            browser.close()


def main():
    # Caspio's "next" occasionally advances two pages at once, so a single pass drops
    # a few pages — but a *different* few each time. Rather than retry from scratch,
    # we UNION successive passes by Code: the gaps are filled within a couple of
    # passes and the merged set converges to the full result.
    merged, total = {}, None
    complete = False
    no_progress = 0
    for attempt in range(1, MAX_ATTEMPTS + 1):
        print(f"\n=== Pass {attempt}/{MAX_ATTEMPTS} "
              f"(have {len(merged)} unique so far) ===")
        before = len(merged)
        try:
            records, pass_total, _ = scrape_once()
        except Exception as exc:      # a flaky pass (e.g. page size wouldn't set)
            print(f"Pass {attempt} failed: {exc}", file=sys.stderr)
            continue
        total = pass_total or total
        for r in records:
            if r["Code"]:
                merged.setdefault(r["Code"], r)
        added = len(merged) - before
        print(f"Pass {attempt}: +{len(records)} rows, {added} new → "
              f"{len(merged)} unique (target {total}).")
        write_csv(list(merged.values()), announce=False)     # checkpoint the union

        if MAX_PAGES:
            complete = True
            break
        if total and len(merged) >= total:
            complete = True
            print(f"Union reached the full set ({len(merged)}/{total}).")
            break
        # If a full pass finds nothing new, the union has converged even if the
        # (unreliable) reported total never matches.
        no_progress = no_progress + 1 if added == 0 else 0
        if attempt >= 2 and no_progress >= 1:
            complete = True
            print(f"Union converged at {len(merged)} (a full pass added nothing new).")
            break

    records = list(merged.values())
    print(f"\nParsed {len(records)} unique records (final total {total}).")
    write_csv(records)

    if not complete:
        print(f"ERROR: union still short after {MAX_ATTEMPTS} passes "
              f"(have {len(records)}, expected ~{total})", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
