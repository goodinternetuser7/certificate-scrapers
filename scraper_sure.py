#!/usr/bin/env python3
"""
Scrapes active (valid) SURE verifications from certification.sure-system.org
Output columns: Client Name, Scope, Issuing CB, Expiry Date, Country
"""

import csv
import sys
import time
from datetime import datetime, timezone

import requests
from bs4 import BeautifulSoup

BASE_URL   = "https://certification.sure-system.org"
SEARCH_URL = f"{BASE_URL}/SearchVerifications"
DELAY      = 1.0   # polite pause between requests
STATUS_VALID = "1"


# ── ASP.NET helpers ──────────────────────────────────────────────────────────

def extract_viewstate(soup):
    def val(id_):
        tag = soup.find("input", {"id": id_})
        return tag["value"] if tag else ""
    return {
        "__VIEWSTATE":          val("__VIEWSTATE"),
        "__VIEWSTATEGENERATOR": val("__VIEWSTATEGENERATOR"),
        "__VIEWSTATEENCRYPTED": val("__VIEWSTATEENCRYPTED"),
        "__EVENTVALIDATION":    val("__EVENTVALIDATION"),
    }

def build_form(soup, event_target="", event_arg=""):
    return {
        **extract_viewstate(soup),
        "__EVENTTARGET":    event_target,
        "__EVENTARGUMENT":  event_arg,
        "ctl00$ContentPlaceHolder$IdentifierTextBox":              "",
        "ctl00$ContentPlaceHolder$CompanyNameTextBox":             "",
        "ctl00$ContentPlaceHolder$CountryDropDownList":            "",
        "ctl00$ContentPlaceHolder$StatusValuesDropDownList":       STATUS_VALID,
        "ctl00$ContentPlaceHolder$ValidFrom$DatePickerTextBox":    "",
        "ctl00$ContentPlaceHolder$ValidTo$DatePickerTextBox":      "",
        "ctl00$ContentPlaceHolder$ScopeTextBox":                   "",
        "ctl00$ContentPlaceHolder$BiomassTextBox":                 "",
        "ctl00$ContentPlaceHolder$CertificationBodyTextBox":       "",
    }

def post(session, soup, event_target="", event_arg=""):
    form = build_form(soup, event_target, event_arg)
    r = session.post(SEARCH_URL, data=form, timeout=60)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser")


# ── Parsing ──────────────────────────────────────────────────────────────────

def get_total(soup):
    span = soup.find("span", {"id": "ContentPlaceHolder_PaginationControl_TotalNumberOfResults"})
    if not span:
        return 0
    try:
        return int(span.text.split()[0].replace(",", ""))
    except ValueError:
        return 0

def get_page_buttons(soup):
    """Returns list of dicts with keys 'value' and 'name' for all paging buttons."""
    panel = soup.find("div", {"id": "ContentPlaceHolder_PaginationControl_RepeaterPanel"})
    if not panel:
        return []
    return [{"value": b.get("value", ""), "name": b.get("name", "")}
            for b in panel.find_all("input", type="button")]

def get_active_page(soup):
    btn = soup.find("input", {"class": "paging-chosen-page"})
    if btn:
        try:
            return int(btn.get("value", ""))
        except ValueError:
            pass
    return None

def parse_rows(soup):
    records = []
    table = soup.find("table", {"id": "ContentPlaceHolder_VerificationsGridView"})
    if not table:
        return records
    for row in table.find_all("tr")[1:]:   # skip header row
        cells = row.find_all("td")
        if len(cells) < 9:
            continue
        # col indices: 0=icon, 1=status, 2=identifier, 3=company+country,
        #              4=valid_from, 5=valid_to, 6=scope, 7=biomass, 8=CB, 9=action
        company_country = cells[3].get_text(strip=True)
        parts      = company_country.rsplit(",", 1)
        client     = parts[0].strip() if len(parts) == 2 else company_country
        country    = parts[1].strip() if len(parts) == 2 else ""
        expiry     = cells[5].get_text(strip=True)
        scope      = cells[6].get_text(strip=True)
        cb         = cells[8].get_text(strip=True)
        records.append({
            "Client Name": client,
            "Scope":       scope,
            "Issuing CB":  cb,
            "Expiry Date": expiry,
            "Country":     country,
        })
    return records


# ── Main scraper ─────────────────────────────────────────────────────────────

def main():
    session = requests.Session()
    session.headers["User-Agent"] = "SURE-cert-scraper/1.0"

    print("Loading page …")
    soup = BeautifulSoup(session.get(SEARCH_URL, timeout=30).text, "html.parser")

    print("Applying 'valid' filter …")
    form = build_form(soup)
    form["ctl00$ContentPlaceHolder$SearchControl$SearchButton.x"] = "0"
    form["ctl00$ContentPlaceHolder$SearchControl$SearchButton.y"] = "0"
    soup = BeautifulSoup(session.post(SEARCH_URL, data=form, timeout=60).text, "html.parser")

    total = get_total(soup)
    print(f"Valid certificates: {total}")

    print("Setting page size to 50 …")
    soup = post(session, soup,
                "ctl00$ContentPlaceHolder$PaginationControl$NumberOfPageResultsLarge")

    # Page 1 is now loaded
    records     = parse_rows(soup)
    visited     = {get_active_page(soup) or 1}

    while True:
        buttons  = get_page_buttons(soup)
        numbered = {int(b["value"]): b["name"]
                    for b in buttons if b["value"].isdigit()}
        next_btn = next((b for b in buttons if b["value"] == "next-page"), None)

        # Scrape every visible numbered page not yet visited, in order
        to_visit = sorted(p for p in numbered if p not in visited)
        for page_num in to_visit:
            print(f"  Page {page_num} …", end="\r", flush=True)
            try:
                soup = post(session, soup, numbered[page_num])
                records.extend(parse_rows(soup))
                visited.add(page_num)
            except requests.RequestException as exc:
                print(f"\nWarning: page {page_num} failed ({exc}), skipping.",
                      file=sys.stderr)
            time.sleep(DELAY)

        if next_btn:
            # Advance to next page group
            soup = post(session, soup, next_btn["name"])
            active = get_active_page(soup)
            if active and active not in visited:
                records.extend(parse_rows(soup))
                visited.add(active)
            time.sleep(DELAY)
        else:
            break  # no more pages

    print(f"\nParsed {len(records)} records.")

    date_str   = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    dated_file = f"SURE certificates {date_str}.csv"
    latest_file = "SURE certificates latest.csv"

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
