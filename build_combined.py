#!/usr/bin/env python3
"""
Builds a single combined workbook from every scheme's committed
`<Scheme> certificates latest.xlsx` dashboard.

Two files are produced (both in one openpyxl session, so the dashboard charts are
never lost to a reload):

  • "All certificates latest.xlsx" — the full workbook:
      - Dashboard    interactive front page (Select Country → pie of certificates
                     by scheme; Select Scheme → bar of top countries), same design
                     as the per-scheme dashboards but driven by the combined set.
      - Data         one normalised, filterable row per certificate across all
                     schemes (Scheme, Identifier, Name, Country, Type,
                     Certification Body, Status, Valid From, Valid To).
      - Certification Bodies
                     filterable table of each certification body → record count
                     → which schemes report it (ISCC, SURE, GGL, SBP publish a
                     CB; PEFC and FSC do not).
      - ISCC … SBP   full native columns per scheme, so no detail is lost.
      - Summary      record counts per scheme.
  • "All certificates (dashboard) latest.xlsx" — the same but without the six
    per-scheme detail sheets, so it's small enough (~10 MB) to email.

The combined set's second dashboard dimension is Scheme (Certification Body is
absent for FSC/PEFC — 89% of rows — so it makes a poor cross-scheme chart split;
the "Certification Bodies" sheet gives the CB breakdown instead).
"""

import csv
import glob
import os
import sys
from datetime import date, datetime, timezone

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from generate_excel import aggregate, build_excel, hdr

SCHEMES = ["ISCC", "SURE", "PEFC", "FSC", "GGL", "SBP"]
BLUE, WHITE = "005798", "FFFFFF"

COMMON = ["Scheme", "Identifier", "Name", "Country", "Type",
          "Certification Body", "Status", "Valid From", "Valid To"]
COMMON_WIDTHS = (10, 20, 42, 18, 26, 26, 12, 13, 13)
COUNTRY_I, SCHEME_I = COMMON.index("Country"), COMMON.index("Scheme")
CB_I = COMMON.index("Certification Body")
DATE_COLS = {"Valid From", "Valid To"}
MAPPINGS = {
    "ISCC": {"Name": "Client Name", "Country": "Country", "Type": "Scope",
             "Certification Body": "Issuing CB", "Valid To": "Expiry Date"},
    "SURE": {"Name": "Client Name", "Country": "Country", "Type": "Scope",
             "Certification Body": "Issuing CB", "Valid To": "Expiry Date"},
    "PEFC": {"Identifier": ("Certificate Number", "Code"), "Name": "Entity",
             "Country": "Country", "Type": "Role", "Status": "Status"},
    "FSC": {"Identifier": "Certificate Code", "Name": "Organization",
            "Country": "Country", "Type": "Certificate Type", "Status": "Status",
            "Valid From": "Valid From", "Valid To": "Valid To"},
    "GGL": {"Identifier": "USI", "Name": "Participant name", "Country": "Country",
            "Type": "Participant role", "Certification Body": "CB",
            "Status": "Status", "Valid From": "Valid from", "Valid To": "Valid till"},
    "SBP": {"Identifier": "Certificate Number", "Name": "Certificate Holder",
            "Country": "Country", "Type": "Certificate Type",
            "Certification Body": "Certification Body", "Status": "Status",
            "Valid From": "Date of Issue", "Valid To": "Date of Expiry"},
}
DATE_FORMATS = ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d %B %Y", "%d %b %Y")

FULL_OUT = "All certificates latest.xlsx"
SLIM_OUT = "All certificates (dashboard) latest.xlsx"
COMBINED_CSV = "All certificates latest.csv"    # transient, feeds the Data sheet


def iso_date(v):
    """Normalise a scheme's date to an ISO 'YYYY-MM-DD' string (sorts uniformly);
    keep the raw value if it matches no known format."""
    if not v:
        return ""
    if isinstance(v, (datetime, date)):
        return (v.date() if isinstance(v, datetime) else v).isoformat()
    s = str(v).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s


def read_data(scheme):
    wb = load_workbook(f"{scheme} certificates latest.xlsx", read_only=True, data_only=True)
    try:
        ws = wb["Data"]
        it = ws.iter_rows(values_only=True)
        headers = list(next(it))
        rows = [["" if c is None else c for c in r]
                for r in it if r and not all(c is None or c == "" for c in r)]
        return headers, rows
    finally:
        wb.close()


def pick(row_map, source):
    if isinstance(source, tuple):
        return next((row_map[c] for c in source if row_map.get(c)), "")
    return row_map.get(source, "")


def add_detail_sheet(wb, scheme, headers, rows):
    ws = wb.create_sheet(scheme)
    ws.append(headers)
    for row in rows:
        ws.append(row)
    for c in range(1, len(headers) + 1):
        hdr(ws.cell(row=1, column=c))
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = min(max(len(str(h)) + 2, 12), 48)
    ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
    tab = Table(displayName=f"{scheme}Data", ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)
    ws.freeze_panes = "A2"


def add_summary_sheet(wb, per_scheme, total):
    ws = wb.create_sheet("Summary")
    ws.append(["Scheme", "Records"])
    for c in (1, 2):
        hdr(ws.cell(row=1, column=c))
    for scheme in SCHEMES:
        if scheme in per_scheme:
            ws.append([scheme, len(per_scheme[scheme][1])])
    ws.append(["TOTAL", total])
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    note_row = ws.max_row + 2
    ws.cell(row=note_row, column=1).value = (
        "Certification Body is published only by ISCC, SURE, GGL and SBP "
        "(see the Certification Bodies sheet). PEFC and FSC do not publish it."
    )
    ws.cell(row=note_row, column=1).font = Font(italic=True, size=9, color="888888")


def add_cb_summary_sheet(wb, combined):
    """One row per certification body: record count + which schemes report it.
    PEFC/FSC carry no CB, so those rows are simply excluded (blank CB)."""
    from collections import defaultdict
    counts = defaultdict(int)
    schemes = defaultdict(set)
    for r in combined:
        cb = r[CB_I]
        cb = cb.strip() if isinstance(cb, str) else cb
        if not cb:
            continue
        counts[cb] += 1
        schemes[cb].add(r[SCHEME_I])

    ws = wb.create_sheet("Certification Bodies")
    ws.append(["Certification Body", "Records", "Schemes"])
    for c in (1, 2, 3):
        hdr(ws.cell(row=1, column=c))
    ordered = sorted(counts.items(), key=lambda x: (-x[1], x[0].lower()))
    for cb, cnt in ordered:
        ws.append([cb, cnt, ", ".join(sorted(schemes[cb]))])
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 24
    ws.freeze_panes = "A2"
    last = len(ordered) + 1
    tab = Table(displayName="CertBodies", ref=f"A1:C{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)
    return len(ordered)


def main():
    combined, per_scheme = [], {}
    for scheme in SCHEMES:
        if not glob.glob(f"{scheme} certificates latest.xlsx"):
            print(f"  skip {scheme}: dashboard not found")
            continue
        headers, rows = read_data(scheme)
        per_scheme[scheme] = (headers, rows)
        m = MAPPINGS[scheme]
        for r in rows:
            row_map = dict(zip(headers, r))
            rec = [scheme]
            for col in COMMON[1:]:
                if col not in m:
                    rec.append("")
                elif col in DATE_COLS:
                    rec.append(iso_date(pick(row_map, m[col])))
                else:
                    rec.append(pick(row_map, m[col]))
            combined.append(rec)
        print(f"  {scheme}: {len(rows)} rows")

    if not combined:
        sys.exit("No dashboards found — did the scrapers run?")

    # Feed the Data sheet + dashboard aggregation. dim2 ('cb' slot) = Scheme.
    with open(COMBINED_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(COMMON)
        w.writerows(combined)
    dash_rows = [{"country": r[COUNTRY_I] or "Unknown", "cb": r[SCHEME_I]} for r in combined]
    country_totals, cb_totals, cb_by_country, country_by_cb = aggregate(dash_rows)
    print(f"Building dashboard: {len(combined)} rows, {len(country_totals)} countries …")

    wb = build_excel(
        dash_rows, country_totals, cb_totals, cb_by_country, country_by_cb, COMBINED_CSV,
        title="All Certificates — Interactive Dashboard",
        dim2_singular="Scheme", dim2_short="Scheme",
        data_fieldnames=COMMON, data_widths=COMMON_WIDTHS,
        kpi_total_label="Total Certificate Records",
        default_prefix="All certificates", save=False,
    )
    n_cbs = add_cb_summary_sheet(wb, combined)
    add_summary_sheet(wb, per_scheme, len(combined))
    print(f"  Certification Bodies sheet: {n_cbs} bodies")

    # Slim (dashboard + data only) first — small enough to email.
    wb.save(SLIM_OUT)
    print(f"Saved → {SLIM_OUT}")

    # Then add per-scheme detail and save the full workbook (+ dated copy).
    for scheme in SCHEMES:
        if scheme in per_scheme:
            add_detail_sheet(wb, scheme, *per_scheme[scheme])
    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"All certificates {date_str}.xlsx", FULL_OUT):
        wb.save(path)
        print(f"Saved → {path}")

    os.remove(COMBINED_CSV)


if __name__ == "__main__":
    main()
