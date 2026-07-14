#!/usr/bin/env python3
"""
Scrapes the Green Gold Label (GGL) certificate holder list.

GGL doesn't expose a search API or an HTML register — instead it publishes the
full holder list as a **PDF** (exported from Excel) linked from its certification
page: https://greengoldlabel.com/certification/ . This scraper finds the most
recent "GGL certificate holder list" PDF linked there, downloads it, and parses
the table into a CSV.

The PDF is a plain multi-page table, but long cells (participant name, role, CB)
wrap across several lines. Each real row is anchored by a numeric USI in the
left column; lines with no USI are wrapped continuations of the row above. We
assign every word to a column by its x-position (the layout is a fixed Excel
grid) and merge continuation lines back into their anchor row.

Output columns:
  USI, Participant name, Country, Participant role, Regulation, Standards,
  Type of biomass, Valid from, Valid till, CB, Status

Unlike the other schemes' registers this list includes *all* statuses (Valid,
Suspended, Withdrawn, Terminated, Expired), so we keep the Status column rather
than pre-filtering — downstream can filter on it.
"""

import csv
import re
import tempfile
from datetime import datetime, timezone

import pdfplumber
import requests

CERT_PAGE = "https://greengoldlabel.com/certification/"
PDF_HREF_RE = re.compile(
    r'href="(https://greengoldlabel\.com/wp-content/uploads/\d{4}/\d{2}/'
    r'GGL[- ]certificate[- ]holder[- ]list[^"]*?\.pdf)"',
    re.I,
)
USER_AGENT = "Mozilla/5.0 GGL-cert-scraper/1.0"

FIELDNAMES = [
    "USI", "Participant name", "Country", "Participant role", "Regulation",
    "Standards", "Type of biomass", "Valid from", "Valid till", "CB", "Status",
]

# Column left/right x-boundaries derived from the PDF header row (A4 landscape,
# 842pt wide). The last column is open-ended. Boundaries are midpoints between
# adjacent header labels, so a word is binned by its horizontal centre.
COL_BOUNDS = [
    (0, 140, "USI"),
    (140, 261, "Participant name"),
    (261, 337, "Country"),
    (337, 420, "Participant role"),
    (420, 474, "Regulation"),
    (474, 543, "Standards"),
    (543, 585, "Type of biomass"),
    (585, 645, "Valid from"),
    (645, 700, "Valid till"),
    (700, 765, "CB"),
    (765, 9999, "Status"),
]
COLS = [c for _, _, c in COL_BOUNDS]

USI_ANCHOR_RE = re.compile(r"^(\d{15,})")   # a real row anchor (18-19 digit USI)
# Per-page footer: a "DD/MM/YYYY" date, optionally trailed by the page number.
# (A bare page-number line lands in the Status column and is dropped by the
# continuation guard, so it needs no rule here — and a rule for it would wrongly
# swallow wrapped values like a lone "5" from "Cat 5".)
FOOTER_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}(?:\s+\d+)?$")
STATUSES = ("Valid", "Suspended", "Withdrawn", "Terminated", "Expired")
FOOTER_DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")  # per-page footer date
MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12,
}


def _col_of(x):
    for lo, hi, name in COL_BOUNDS:
        if lo <= x < hi:
            return name
    return "Status"


def _pdf_date(url):
    """Sortable date for a holder-list URL, for picking the newest link. Uses the
    reliable /uploads/YYYY/MM/ path (always present per PDF_HREF_RE) as the base,
    refined to the day from the filename's 'DD-Month-YYYY' when it parses."""
    ym = re.search(r"/uploads/(\d{4})/(\d{2})/", url)
    year, month, day = (int(ym.group(1)), int(ym.group(2)), 1) if ym else (1970, 1, 1)
    m = re.search(r"(\d{1,2})[- ]([A-Za-z]+)[- ](\d{4})\.pdf$", url)
    if m and m.group(2).lower() in MONTHS:
        year, month, day = int(m.group(3)), MONTHS[m.group(2).lower()], int(m.group(1))
    return datetime(year, month, day)


def find_latest_pdf_url(session):
    r = session.get(CERT_PAGE, timeout=60)
    r.raise_for_status()
    urls = list(dict.fromkeys(PDF_HREF_RE.findall(r.text)))
    if not urls:
        raise SystemExit(f"No holder-list PDF link found on {CERT_PAGE}")
    urls.sort(key=_pdf_date)
    return urls[-1]


def parse_pdf(path):
    records = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            # Cluster words into visual lines by their (rounded) top coordinate.
            lines = {}
            for w in page.extract_words(use_text_flow=False, keep_blank_chars=False):
                lines.setdefault(round(w["top"] / 3) * 3, []).append(w)

            for key in sorted(lines):
                buckets = {c: [] for c in COLS}
                for w in sorted(lines[key], key=lambda w: w["x0"]):
                    buckets[_col_of((w["x0"] + w["x1"]) / 2)].append(w["text"])
                cell = {c: " ".join(buckets[c]).strip() for c in COLS}
                joined = " ".join(cell.values()).strip()

                if not joined:
                    continue
                # Skip repeated page furniture: title, the two header lines
                # ("USI …" and its "biomass" wrap), and the footer date/page no.
                if "GGL Certificate Holder List" in joined:
                    continue
                if cell["USI"] == "USI":
                    continue
                if joined == "biomass":
                    continue
                if FOOTER_RE.match(joined):
                    continue

                anchor = USI_ANCHOR_RE.match(cell["USI"])
                if anchor:
                    cell["USI"] = anchor.group(1)   # drop any stray trailing chars
                    records.append(cell)
                elif records and not (cell["USI"] or cell["Valid from"]
                                      or cell["Valid till"] or cell["Status"]):
                    # A genuine wrapped continuation carries no USI, dates, or
                    # status — only overflow text in name/role/standards/CB. This
                    # guard stops stray page furniture merging into a real row.
                    for c in COLS:
                        if cell[c]:
                            records[-1][c] = (records[-1][c] + " " + cell[c]).strip()

    # Occasionally the Excel export glues the CB and Status cells into one word
    # (e.g. "BM CertificationValid" with no space), so the status lands in the CB
    # column and Status is left empty. Split the trailing status keyword back out.
    for rec in records:
        if not rec["Status"]:
            for st in STATUSES:
                # Only split genuine glue ("…CertificationValid"): the status
                # suffix must follow a lowercase letter (a word-boundary run-on),
                # never a real CB name that merely ends in a capitalised word.
                if (rec["CB"].endswith(st) and len(rec["CB"]) > len(st)
                        and rec["CB"][-len(st) - 1].islower()):
                    rec["CB"] = rec["CB"][: -len(st)].strip()
                    rec["Status"] = st
                    break
    return records


def main():
    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT

    print(f"Finding latest holder-list PDF on {CERT_PAGE} …")
    pdf_url = find_latest_pdf_url(session)
    print(f"  → {pdf_url}")

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        resp = session.get(pdf_url, timeout=120)
        resp.raise_for_status()
        tmp.write(resp.content)
        tmp_path = tmp.name
    print(f"Downloaded {len(resp.content):,} bytes. Parsing …")

    records = parse_pdf(tmp_path)
    if not records:
        raise SystemExit("Parsed 0 records — PDF layout may have changed.")
    valid = sum(1 for r in records if r["Status"].lower() == "valid")
    print(f"Parsed {len(records)} holders ({valid} Valid).")

    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"GGL certificates {date_str}.csv", "GGL certificates latest.csv"):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDNAMES)
            w.writeheader()
            w.writerows(records)
        print(f"Saved → {path}")


if __name__ == "__main__":
    main()
