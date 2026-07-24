#!/usr/bin/env python3
"""
Generates the interactive Excel dashboard for RSPO members (scraped from the
member register behind https://rspo.org/search-members/ — see scraper_rspo.py).
The second dashboard dimension is Sector. Reuses aggregate() and build_excel()
from generate_excel.py.
"""
import csv
import glob

from generate_excel import aggregate, build_excel

PREFIX = "RSPO members"
DATA_FIELDS = ("Membership Number", "Member Name", "Country",
               "Membership Category", "Sector", "Status", "Last Update",
               "Group Members", "Group Member Names", "Profile URL")
DATA_WIDTHS = (20, 44, 20, 20, 34, 12, 14, 14, 60, 34)


def find_csv():
    dated = sorted(glob.glob(f"{PREFIX} 20*.csv"))
    return dated[-1] if dated else f"{PREFIX} latest.csv"


def load_data(path):
    """Country + Sector (the member's RSPO membership sector)."""
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            country = (r.get("Country") or "Unknown").strip() or "Unknown"
            sector = (r.get("Sector") or "Unknown").strip() or "Unknown"
            rows.append({"country": country, "cb": sector})
    return rows


def main():
    csv_path = find_csv()
    print(f"Reading {csv_path} …")
    rows = load_data(csv_path)
    print(f"Loaded {len(rows)} records. Aggregating …")
    country_totals, sector_totals, sector_by_country, country_by_sector = aggregate(rows)
    print(f"  {len(country_totals)} countries  |  {len(sector_totals)} sectors")

    date_part = csv_path.replace(f"{PREFIX} ", "").replace(".csv", "")
    build_excel(
        rows, country_totals, sector_totals, sector_by_country, country_by_sector,
        csv_path,
        title="RSPO Members — Interactive Dashboard",
        dated_out=f"{PREFIX} {date_part}.xlsx",
        latest_out=f"{PREFIX} latest.xlsx",
        dim2_singular="Sector", dim2_short="Sector",
        data_fieldnames=DATA_FIELDS, data_widths=DATA_WIDTHS,
        kpi_total_label="Total Members",
        default_prefix=PREFIX,
    )


if __name__ == "__main__":
    main()
