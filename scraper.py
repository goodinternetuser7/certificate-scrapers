#!/usr/bin/env python3
"""
Scrapes active (valid) ISCC certificates from iscc-system.org and writes a CSV with:
Client Name, Scope, Issuing CB, Expiry Date, Country
"""

import csv
import json
import sys
import time
from datetime import datetime, timezone

import requests
from bs4 import BeautifulSoup

API_URL = "https://iscc-system.org/wp-json/api/certificates"
PAGE_SIZE = 200
REQUEST_DELAY = 0.5  # seconds between pages


def fetch_page(session: requests.Session, page: int) -> dict:
    payload = {
        "filters": {"status": ["valid"]},
        "count": PAGE_SIZE,
        "page": page,
    }
    r = session.post(API_URL, json=payload, timeout=60)
    r.raise_for_status()
    # API response may include a UTF-8 BOM
    data = json.loads(r.content.decode("utf-8-sig"))
    return data["data"]["data"]


def parse_cards(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    records = []

    for card in soup.find_all("div", class_="is-certificate"):
        # Client name and country — visible text format: "Company,  City[, Region], Country"
        # The double-space after the comma marks the boundary between company and location.
        name_el = card.select_one("h3.h4 span")
        raw_name = name_el.get_text() if name_el else ""
        client_name = raw_name.strip()
        country = ""
        if ",  " in raw_name:
            location_part = raw_name.split(",  ", 1)[1]
            country = location_part.split(",")[-1].strip()
        elif client_name:
            # Fallback: last comma-separated token
            country = client_name.split(",")[-1].strip()

        # Validity period "DD.MM.YY – DD.MM.YY" — take the end date
        date_el = card.select_one("div.date")
        expiry = ""
        if date_el:
            parts = date_el.get_text(strip=True).split("–")
            if len(parts) == 2:
                raw = parts[1].strip()
                # Convert DD.MM.YY → DD.MM.20YY for clarity
                try:
                    d, m, y = raw.split(".")
                    expiry = f"{d}.{m}.20{y}"
                except ValueError:
                    expiry = raw

        # Fold items: Scope, Issuing CB
        scope = issuing_cb = ""
        for item in card.select("div.is-certificate-fold-item"):
            label_el = item.select_one("p.title")
            value_els = item.find_all("p")
            if not label_el or len(value_els) < 2:
                continue
            label = label_el.get_text(strip=True)
            value = value_els[-1].get_text(strip=True)
            if label == "Scope":
                scope = value
            elif label == "Issuing CB":
                issuing_cb = value

        records.append({
            "Client Name": client_name,
            "Scope": scope,
            "Issuing CB": issuing_cb,
            "Expiry Date": expiry,
            "Country": country,
        })

    return records


def main():
    session = requests.Session()
    session.headers.update({"User-Agent": "ISCC-cert-scraper/1.0"})

    print("Fetching page 1 …")
    first = fetch_page(session, 1)
    max_pages = first["maxPages"]
    total = first["totalCount"]
    print(f"Active certificates: {total}  |  Pages to fetch: {max_pages}")

    records = parse_cards(first["html"])

    for page in range(2, max_pages + 1):
        print(f"  Page {page}/{max_pages} …", end="\r", flush=True)
        try:
            data = fetch_page(session, page)
            records.extend(parse_cards(data["html"]))
        except requests.RequestException as exc:
            print(f"\nWarning: page {page} failed ({exc}), skipping.", file=sys.stderr)
        time.sleep(REQUEST_DELAY)

    print(f"\nParsed {len(records)} records.")

    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    dated_file = f"certificates_{date_str}.csv"
    latest_file = "certificates_latest.csv"

    fieldnames = ["Client Name", "Scope", "Issuing CB", "Expiry Date", "Country"]
    for path in (dated_file, latest_file):
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(records)

    print(f"Saved → {dated_file}")
    print(f"Saved → {latest_file}")


if __name__ == "__main__":
    main()
