#!/usr/bin/env python3
"""
Generates the interactive Excel dashboard for GLOBALG.A.P. producers
(scraped from the FoodPLUS/osapiens Supply Chain Portal, currently limited to
the Baltic countries — see scraper_ggap.py). The second dashboard dimension is
Product. Reuses aggregate() and build_excel() from generate_excel.py.
"""
import csv
import glob

from generate_excel import aggregate, build_excel

PREFIX = "GGAP certificates"
DATA_FIELDS = ("GGN", "Producer Name", "City", "Country", "Producer Type", "Product")
DATA_WIDTHS = (16, 44, 24, 14, 16, 26)


def find_csv():
    dated = sorted(glob.glob(f"{PREFIX} 20*.csv"))
    return dated[-1] if dated else f"{PREFIX} latest.csv"


def load_data(path):
    """Country + Product (the crop the producer is certified for)."""
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            country = (r.get("Country") or "Unknown").strip() or "Unknown"
            product = (r.get("Product") or "Unknown").strip() or "Unknown"
            rows.append({"country": country, "cb": product})
    return rows


def main():
    csv_path = find_csv()
    print(f"Reading {csv_path} …")
    rows = load_data(csv_path)
    print(f"Loaded {len(rows)} records. Aggregating …")
    country_totals, prod_totals, prod_by_country, country_by_prod = aggregate(rows)
    print(f"  {len(country_totals)} countries  |  {len(prod_totals)} products")

    date_part = csv_path.replace(f"{PREFIX} ", "").replace(".csv", "")
    build_excel(
        rows, country_totals, prod_totals, prod_by_country, country_by_prod, csv_path,
        title="GLOBALG.A.P. Producers (Baltics) — Interactive Dashboard",
        dated_out=f"{PREFIX} {date_part}.xlsx",
        latest_out=f"{PREFIX} latest.xlsx",
        dim2_singular="Product", dim2_short="Product",
        data_fieldnames=DATA_FIELDS, data_widths=DATA_WIDTHS,
        kpi_total_label="Total Producer Records",
        default_prefix=PREFIX,
    )


if __name__ == "__main__":
    main()
