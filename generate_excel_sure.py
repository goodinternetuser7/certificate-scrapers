#!/usr/bin/env python3
"""
Generates the interactive Excel dashboard for SURE certificates.
Reuses the same build_excel() logic as generate_excel.py.
"""
import glob
from generate_excel import load_data, aggregate, build_excel

PREFIX = "SURE certificates"

def find_csv():
    dated = sorted(glob.glob(f"{PREFIX} 20*.csv"))
    return dated[-1] if dated else f"{PREFIX} latest.csv"

def main():
    csv_path = find_csv()
    print(f"Reading {csv_path} …")
    rows = load_data(csv_path)
    print(f"Loaded {len(rows)} records. Aggregating …")
    country_totals, cb_totals, cb_by_country, country_by_cb = aggregate(rows)
    print(f"  {len(country_totals)} countries  |  {len(cb_totals)} CBs")

    # Patch the file-naming so build_excel produces SURE-named files
    date_part = csv_path.replace(f"{PREFIX} ", "").replace(".csv", "")
    build_excel(
        rows, country_totals, cb_totals, cb_by_country, country_by_cb,
        csv_path,
        title="SURE Active Certificates — Interactive Dashboard",
        dated_out=f"{PREFIX} {date_part}.xlsx",
        latest_out=f"{PREFIX} latest.xlsx",
    )

if __name__ == "__main__":
    main()
