# Certificate Scrapers

Monthly scrapers that fetch **active / valid** certificates from public certification
registries and commit an interactive Excel dashboard (plus a dated archive) back to
this repo via GitHub Actions.

| Scheme | Source | Scraper | Dashboard | Workflow (cron, UTC 1st) |
|---|---|---|---|---|
| **ISCC** | [iscc-system.org](https://iscc-system.org/certification/all-certificates/) | `scraper.py` | `generate_excel.py` | `monthly-scrape.yml` (06:00) |
| **SURE** | [certification.sure-system.org](https://certification.sure-system.org/SearchVerifications) | `scraper_sure.py` | `generate_excel_sure.py` | `monthly-scrape-sure.yml` (07:00) |
| **PEFC** | [pefc.org/find-certified-legacy](https://pefc.org/find-certified-legacy) | `scraper_pefc.py` | `generate_excel_pefc.py` | `monthly-scrape-pefc.yml` (08:00) |

Each run produces `<Scheme> certificates latest.xlsx` (most recent) and a dated
`<Scheme> certificates YYYY.MM.DD.xlsx` archive. You can also trigger any scraper
manually from the **Actions** tab → *Run workflow*.

## ISCC & SURE

Pure-HTTP scrapers (ISCC via a JSON API, SURE via an ASP.NET search form). Columns:
`Client Name, Scope, Issuing CB, Expiry Date, Country`.

```bash
pip install -r requirements.txt
python scraper.py        # or scraper_sure.py
python generate_excel.py # or generate_excel_sure.py
```

## PEFC

PEFC's "Find certified" search is a **Caspio DataPage** that only renders in a
browser, so `scraper_pefc.py` drives headless Chromium via **Playwright**: it filters
Status = Valid, sets 250 results per page, and pages through the whole set
(~50k records). Caspio's pagination is unreliable (a "next" click occasionally
skips a page), so the scraper makes several passes and **unions the results by
record code** until the merged set stops growing — so the final dataset is
complete even though any single pass isn't. Expect a few passes (~30 min each).
Columns:

| Column | Description |
|---|---|
| Code | PEFC internal record code |
| Entity | Organisation name |
| City | City / postal line |
| Country | Country |
| Role | Relationship (e.g. *Certificate holder*, *Site/Member*) |
| Certificate Number | CoC / FM certificate number (if any) |
| Licence Number | PEFC logo licence number (if any) |
| Category | e.g. *COC - Multisite*, *D - Other* |
| Status | Certificate status (Valid) |
| Type | *Certificate* or *Licence* |
| Entity ID / Certificate ID / Licence ID | PEFC EID / CID / LID keys |

> **Note:** the public PEFC list does not expose the issuing Certification Body or
> the expiry date (those live on per-entity detail pages, one request each — not
> feasible for ~66k rows), so those columns are omitted for PEFC.

```bash
pip install -r requirements-pefc.txt
python -m playwright install chromium
python scraper_pefc.py
python generate_excel_pefc.py
```

Handy env vars for local runs: `PEFC_MAX_PAGES=3` (limit pages),
`PEFC_PAGE_SIZE=250`, `PEFC_HEADFUL=1` (visible browser).
