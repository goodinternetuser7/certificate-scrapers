#!/usr/bin/env python3
"""
Generates the interactive Excel dashboard for the ENplus® pellet producer
register (scraped from https://enplus-pellets.eu/producer/ — see
scraper_enplus.py). The second dashboard dimension is Certification Body, which
ENplus publishes for (almost) every company. Reuses aggregate() and build_excel()
from generate_excel.py.
"""
import csv
import glob

from generate_excel import aggregate, build_excel

PREFIX = "ENplus certificates"
DATA_FIELDS = ("ENplus ID", "Producer", "Status", "Status Since", "Country",
               "City", "Certification Body", "Quality Classes",
               "Certified Activities", "Legal Address", "Certified Sites",
               "Certified Site Names", "Bag Designs", "Website", "Company ID")
DATA_WIDTHS = (12, 44, 12, 13, 20, 28, 34, 34, 60, 50, 14, 50, 12, 30, 12)


def find_csv():
    dated = sorted(glob.glob(f"{PREFIX} 20*.csv"))
    return dated[-1] if dated else f"{PREFIX} latest.csv"


def load_data(path):
    """Country + Certification Body."""
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
    print(f"  {len(country_totals)} countries  |  {len(cb_totals)} certification bodies")

    date_part = csv_path.replace(f"{PREFIX} ", "").replace(".csv", "")
    build_excel(
        rows, country_totals, cb_totals, cb_by_country, country_by_cb, csv_path,
        title="ENplus® Certified Producers — Interactive Dashboard",
        dated_out=f"{PREFIX} {date_part}.xlsx",
        latest_out=f"{PREFIX} latest.xlsx",
        dim2_singular="Certification Body", dim2_short="CB",
        data_fieldnames=DATA_FIELDS, data_widths=DATA_WIDTHS,
        kpi_total_label="Total Producers",
        default_prefix=PREFIX,
    )


if __name__ == "__main__":
    main()
