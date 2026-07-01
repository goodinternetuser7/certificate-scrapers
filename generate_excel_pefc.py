#!/usr/bin/env python3
"""
Generates the interactive Excel dashboard for PEFC certificates & licences.

PEFC's public list has no issuing-CB column, so the second dimension is the
certificate/licence Category (e.g. "COC - Multisite", "D - Other") instead of CB.
Reuses aggregate() and build_excel() from generate_excel.py.
"""
import csv
import glob

from generate_excel import aggregate, build_excel

PREFIX = "PEFC certificates"
DATA_FIELDS = ("Code", "Entity", "City", "Country", "Role",
               "Certificate Number", "Licence Number",
               "Category", "Status", "Type",
               "Entity ID", "Certificate ID", "Licence ID")
DATA_WIDTHS = (12, 45, 28, 18, 16, 24, 20, 20, 10, 12, 12, 14, 14)


def find_csv():
    dated = sorted(glob.glob(f"{PREFIX} 20*.csv"))
    return dated[-1] if dated else f"{PREFIX} latest.csv"


def load_data(path):
    """Country + Category (the CB slot is filled with Category for PEFC)."""
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            country = (r.get("Country") or "Unknown").strip() or "Unknown"
            category = (r.get("Category") or "Unknown").strip() or "Unknown"
            rows.append({"country": country, "cb": category})
    return rows


def main():
    csv_path = find_csv()
    print(f"Reading {csv_path} …")
    rows = load_data(csv_path)
    print(f"Loaded {len(rows)} records. Aggregating …")
    country_totals, cat_totals, cat_by_country, country_by_cat = aggregate(rows)
    print(f"  {len(country_totals)} countries  |  {len(cat_totals)} categories")

    date_part = csv_path.replace(f"{PREFIX} ", "").replace(".csv", "")
    build_excel(
        rows, country_totals, cat_totals, cat_by_country, country_by_cat, csv_path,
        title="PEFC Valid Certificates & Licences — Interactive Dashboard",
        dated_out=f"{PREFIX} {date_part}.xlsx",
        latest_out=f"{PREFIX} latest.xlsx",
        dim2_singular="Category", dim2_short="Category",
        data_fieldnames=DATA_FIELDS, data_widths=DATA_WIDTHS,
        kpi_total_label="Total Valid Records",
        default_prefix=PREFIX,
    )


if __name__ == "__main__":
    main()
