#!/usr/bin/env python3
"""
Builds a single workbook combining every scheme's data from the committed
`<Scheme> certificates latest.xlsx` dashboards:

  • "All Certificates" — one normalised, filterable row per certificate across
    all schemes (Scheme, Identifier, Name, Country, Type, Certification Body,
    Status, Valid From, Valid To), with dates normalised to real Excel dates so
    the whole set sorts/filters together.
  • one sheet per scheme — the full native columns, as a filterable table, so no
    detail (products, licence numbers, city, …) is lost.
  • "Summary" — row counts per scheme.

Output: "All certificates latest.xlsx" (+ a dated copy).
"""

import glob
import sys
from datetime import date, datetime, timezone

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

SCHEMES = ["ISCC", "SURE", "PEFC", "FSC", "GGL", "SBP"]

BLUE = "005798"
WHITE = "FFFFFF"

# Common (normalised) schema, and how each scheme's source columns map onto it.
# A tuple value means "first non-empty of these source columns".
COMMON = ["Scheme", "Identifier", "Name", "Country", "Type",
          "Certification Body", "Status", "Valid From", "Valid To"]
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


def norm_date(v):
    """Parse a scheme's date string into a real date for uniform sorting; keep
    the raw value if it doesn't match any known format."""
    if not v:
        return ""
    if isinstance(v, (datetime, date)):
        return v if not isinstance(v, datetime) else v.date()
    s = str(v).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return s


def read_data(scheme):
    """Return (headers, rows) from a scheme's dashboard Data sheet."""
    wb = load_workbook(f"{scheme} certificates latest.xlsx", read_only=True, data_only=True)
    try:
        ws = wb["Data"]
        it = ws.iter_rows(values_only=True)
        headers = list(next(it))
        rows = []
        for r in it:
            if r is None or all(c is None or c == "" for c in r):
                continue
            rows.append(["" if c is None else c for c in r])
        return headers, rows
    finally:
        wb.close()


def pick(row_map, source):
    if isinstance(source, tuple):
        for col in source:
            if row_map.get(col):
                return row_map[col]
        return ""
    return row_map.get(source, "")


def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(color=WHITE, bold=True, name="Calibri", size=11)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def add_table(ws, name, ncols, nrows):
    ref = f"A1:{get_column_letter(ncols)}{nrows + 1}"
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)
    ws.freeze_panes = "A2"


def write_sheet(wb, title, headers, rows, table_name, date_cols=()):
    ws = wb.create_sheet(title)
    ws.append(headers)
    date_idx = [i for i, h in enumerate(headers) if h in date_cols]
    for row in rows:
        ws.append(row)
    style_header(ws, len(headers))
    # Apply date number format to normalised date columns
    for ci in date_idx:
        col = get_column_letter(ci + 1)
        for r in range(2, len(rows) + 2):
            ws[f"{col}{r}"].number_format = "yyyy-mm-dd"
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = min(max(len(str(h)) + 2, 12), 48)
    add_table(ws, table_name, len(headers), len(rows))
    return ws


def main():
    combined = []
    per_scheme = {}
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
                    rec.append(norm_date(pick(row_map, m[col])))
                else:
                    rec.append(pick(row_map, m[col]))
            combined.append(rec)
        print(f"  {scheme}: {len(rows)} rows")

    if not combined:
        sys.exit("No dashboards found — did the scrapers run?")

    wb = Workbook()
    wb.remove(wb.active)

    print(f"Writing All Certificates ({len(combined)} rows) …")
    write_sheet(wb, "All Certificates", COMMON, combined, "AllCertificates", DATE_COLS)

    for scheme in SCHEMES:
        if scheme in per_scheme:
            headers, rows = per_scheme[scheme]
            write_sheet(wb, scheme, headers, rows, f"{scheme}Data")

    ws = wb.create_sheet("Summary")
    ws.append(["Scheme", "Records"])
    style_header(ws, 2)
    for scheme in SCHEMES:
        if scheme in per_scheme:
            ws.append([scheme, len(per_scheme[scheme][1])])
    ws.append(["TOTAL", len(combined)])
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12

    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"All certificates {date_str}.xlsx", "All certificates latest.xlsx"):
        wb.save(path)
        print(f"Saved → {path}")


if __name__ == "__main__":
    main()
