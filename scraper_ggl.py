#!/usr/bin/env python3
"""
Scrapes the Green Gold Label (GGL) certificate-holder register.

GGL replaced its PDF holder-list (previously linked from /certification/) with a
live register at https://greengoldlabel.com/participants/register that publishes
an official structured export: an .xlsx behind /api/register-export. We fetch
that export directly and reshape it into our CSV, which is far more robust than
the old PDF column-binning (no fixed x-position grid, no wrapped-cell merging).

The register page links the export as `registerDownload.href`; we read that href
off the page each run and fall back to the well-known default if it moves.

Output columns (unchanged, so the downstream dashboard/combined pipeline is
untouched):
  USI, Participant name, Country, Participant role, Regulation, Standards,
  Type of biomass, Valid from, Valid till, CB, Status

Unlike the other schemes' registers this list includes *all* statuses (Valid,
Suspended, Expired, Revoked, Resigned, …), so we keep the Status column rather
than pre-filtering — downstream can filter on it.
"""

import csv
import io
import re
import tempfile
from datetime import datetime, timezone

import openpyxl
import requests

REGISTER_PAGE = "https://greengoldlabel.com/participants/register"
DEFAULT_EXPORT = "https://greengoldlabel.com/api/register-export"
EXPORT_HREF_RE = re.compile(r'href="(/api/register-export[^"]*)"', re.I)
BASE = "https://greengoldlabel.com"
USER_AGENT = "Mozilla/5.0 GGL-cert-scraper/2.0"

FIELDNAMES = [
    "USI", "Participant name", "Country", "Participant role", "Regulation",
    "Standards", "Type of biomass", "Valid from", "Valid till", "CB", "Status",
]

# The export's own header labels → our CSV field names. Matched case-insensitively
# on the stripped label, so column order and minor casing shifts don't matter.
SRC_TO_FIELD = {
    "usi": "USI",
    "name": "Participant name",
    "country": "Country",
    "role": "Participant role",
    "regulation": "Regulation",
    "standards": "Standards",
    "biomass": "Type of biomass",
    "valid from": "Valid from",
    "valid till": "Valid till",
    "certification body": "CB",
    "status": "Status",
}


def find_export_url(session):
    """Read the /api/register-export href off the register page, so a path change
    is picked up automatically; fall back to the known default if it's absent."""
    try:
        r = session.get(REGISTER_PAGE, timeout=60)
        r.raise_for_status()
        m = EXPORT_HREF_RE.search(r.text)
        if m:
            return BASE + m.group(1)
    except requests.RequestException as e:
        print(f"  (register page unreadable: {e}; using default export URL)")
    return DEFAULT_EXPORT


def _norm(v):
    """Cell value → trimmed string; render dates as ISO YYYY-MM-DD."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def parse_export(content):
    """Parse the official register .xlsx (bytes) into our CSV rows.

    The sheet carries a few title/meta rows, then a header row starting with the
    'USI' cell, then the data. We locate the header by that 'USI' cell and map
    every column by its label, so we never depend on a fixed column order."""
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    list_date = ""
    header_idx = None
    for i, row in enumerate(rows):
        first = _norm(row[0] if row else "")
        if first.lower().startswith("list date"):
            list_date = first.split(":", 1)[-1].strip()
        if first == "USI":
            header_idx = i
            break
    if header_idx is None:
        raise SystemExit("No 'USI' header row found — register export layout may "
                         "have changed.")
    if list_date:
        print(f"  register list date: {list_date}")

    # column index → our field name (only columns we recognise)
    col_field = {}
    for ci, label in enumerate(rows[header_idx]):
        key = _norm(label).lower()
        if key in SRC_TO_FIELD:
            col_field[ci] = SRC_TO_FIELD[key]
    missing = set(FIELDNAMES) - set(col_field.values())
    if missing:
        raise SystemExit(f"Register export is missing expected columns: {sorted(missing)}")

    records = []
    for row in rows[header_idx + 1:]:
        if not row or not _norm(row[0] if row else ""):
            continue  # skip blank / trailing note rows (a data row always has a USI)
        rec = {f: "" for f in FIELDNAMES}
        for ci, field in col_field.items():
            if ci < len(row):
                rec[field] = _norm(row[ci])
        if rec["USI"]:
            records.append(rec)
    return records


def main():
    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT

    export_url = find_export_url(session)
    print(f"Downloading GGL register export from {export_url} …")
    resp = session.get(export_url, timeout=120)
    resp.raise_for_status()
    print(f"Downloaded {len(resp.content):,} bytes. Parsing …")

    records = parse_export(resp.content)
    if not records:
        raise SystemExit("Parsed 0 records — register export may have changed.")
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
