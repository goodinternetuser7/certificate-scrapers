#!/usr/bin/env python3
"""
Generates the interactive Excel dashboard for FSC certificates.
Second dimension is the certificate Type (COC / FM / CW / …). Reuses
aggregate() and build_excel() from generate_excel.py.
"""
import csv
import glob

from generate_excel import aggregate, build_excel

PREFIX = "FSC certificates"
DATA_FIELDS = ("Licence Code", "Certificate Code", "Certificate Type", "Status",
               "Controlled Wood", "Valid From", "Valid To", "Organization",
               "Role", "Site Status", "State/Province", "Country")
DATA_WIDTHS = (14, 22, 12, 10, 10, 12, 12, 42, 16, 12, 18, 18)


def find_csv():
    dated = sorted(glob.glob(f"{PREFIX} 20*.csv"))
    return dated[-1] if dated else f"{PREFIX} latest.csv"


def load_data(path):
    """Country + Certificate Type (the CB slot holds the type for FSC)."""
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            country = (r.get("Country") or "Unknown").strip() or "Unknown"
            ctype = (r.get("Certificate Type") or "Unknown").strip() or "Unknown"
            rows.append({"country": country, "cb": ctype})
    return rows


def main():
    csv_path = find_csv()
    print(f"Reading {csv_path} …")
    rows = load_data(csv_path)
    print(f"Loaded {len(rows)} records. Aggregating …")
    country_totals, type_totals, type_by_country, country_by_type = aggregate(rows)
    print(f"  {len(country_totals)} countries  |  {len(type_totals)} types")

    date_part = csv_path.replace(f"{PREFIX} ", "").replace(".csv", "")
    build_excel(
        rows, country_totals, type_totals, type_by_country, country_by_type, csv_path,
        title="FSC Valid Certificates — Interactive Dashboard",
        dated_out=f"{PREFIX} {date_part}.xlsx",
        latest_out=f"{PREFIX} latest.xlsx",
        dim2_singular="Certificate Type", dim2_short="Type",
        data_fieldnames=DATA_FIELDS, data_widths=DATA_WIDTHS,
        kpi_total_label="Total Valid Certificates",
        default_prefix=PREFIX,
    )


if __name__ == "__main__":
    main()
