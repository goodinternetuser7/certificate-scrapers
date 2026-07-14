#!/usr/bin/env python3
"""
Generates the interactive Excel dashboard for Sustainable Biomass Program (SBP)
certificate holders. Second dimension is the Certification Body (CB). Reuses
aggregate() and build_excel() from generate_excel.py.
"""
import csv
import glob

from generate_excel import aggregate, build_excel

PREFIX = "SBP certificates"
DATA_FIELDS = ("Certificate Number", "Certificate Holder", "Country",
               "Certificate Type", "Status", "Certification Body",
               "Date of Issue", "Date of Expiry", "Certificate Scope",
               "Products Covered")
DATA_WIDTHS = (16, 40, 18, 18, 12, 30, 16, 16, 30, 44)


def find_csv():
    dated = sorted(glob.glob(f"{PREFIX} 20*.csv"))
    return dated[-1] if dated else f"{PREFIX} latest.csv"


def load_data(path):
    """Country + CB (the issuing Certification Body)."""
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            country = (r.get("Country") or "Unknown").strip() or "Unknown"
            cb = (r.get("Certification Body") or "Unknown").strip() or "Unknown"
            rows.append({"country": country, "cb": cb})
    return rows


def main():
    csv_path = find_csv()
    print(f"Reading {csv_path} …")
    rows = load_data(csv_path)
    print(f"Loaded {len(rows)} records. Aggregating …")
    country_totals, cb_totals, cb_by_country, country_by_cb = aggregate(rows)
    print(f"  {len(country_totals)} countries  |  {len(cb_totals)} CBs")

    date_part = csv_path.replace(f"{PREFIX} ", "").replace(".csv", "")
    build_excel(
        rows, country_totals, cb_totals, cb_by_country, country_by_cb, csv_path,
        title="SBP Certificate Holders — Interactive Dashboard",
        dated_out=f"{PREFIX} {date_part}.xlsx",
        latest_out=f"{PREFIX} latest.xlsx",
        dim2_singular="Certification Body", dim2_short="CB",
        data_fieldnames=DATA_FIELDS, data_widths=DATA_WIDTHS,
        kpi_total_label="Total Certificate Holders",
        default_prefix=PREFIX,
    )


if __name__ == "__main__":
    main()
