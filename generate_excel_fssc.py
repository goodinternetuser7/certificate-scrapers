#!/usr/bin/env python3
"""
Generates the interactive Excel dashboard for FSSC certified organizations
(scraped from the FSSC public register — see scraper_fssc.py). The register
publishes no certification body, so the second dashboard dimension is the food
chain Category. Reuses aggregate() and build_excel() from generate_excel.py.
"""
import csv
import glob

from generate_excel import aggregate, build_excel

PREFIX = "FSSC certificates"
DATA_FIELDS = ("COID", "Organization", "Country", "City", "Scheme", "Status",
               "Food Chain Category", "Product Types", "Scope Statement",
               "Initial Certification", "Issued", "Valid Until",
               "Last Status Decision", "GFSI Recognized")
DATA_WIDTHS = (20, 44, 18, 20, 14, 12, 40, 26, 50, 16, 14, 14, 16, 14)


def find_csv():
    dated = sorted(glob.glob(f"{PREFIX} 20*.csv"))
    return dated[-1] if dated else f"{PREFIX} latest.csv"


def load_data(path):
    """Country + food chain Category (the register publishes no CB).

    A quarter of the organizations are certified for several food chain
    categories; the dashboard keys on the *first* one, so the selector stays the
    register's own list of 16 categories instead of ~200 combination strings.
    The Data sheet still carries every category per row."""
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            country = (r.get("Country") or "Unknown").strip() or "Unknown"
            cats = [c.strip() for c in (r.get("Food Chain Category") or "").split(";")]
            cats = [c for c in cats if c and c != ":"]     # a few rows carry an empty category
            rows.append({"country": country, "cb": cats[0] if cats else "Unknown"})
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
        title="FSSC Certified Organizations — Interactive Dashboard",
        dated_out=f"{PREFIX} {date_part}.xlsx",
        latest_out=f"{PREFIX} latest.xlsx",
        dim2_singular="Category", dim2_short="Category",
        data_fieldnames=DATA_FIELDS, data_widths=DATA_WIDTHS,
        kpi_total_label="Total Certified Organizations",
        default_prefix=PREFIX,
    )


if __name__ == "__main__":
    main()
