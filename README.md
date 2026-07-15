# Certificate Scrapers

Monthly scrapers that fetch **active / valid** certificates from public certification
registries and commit an interactive Excel dashboard (plus a dated archive) back to
this repo via GitHub Actions.

| Scheme | Source | Scraper | Dashboard | Workflow (cron, UTC 1st) |
|---|---|---|---|---|
| **ISCC** | [iscc-system.org](https://iscc-system.org/certification/all-certificates/) | `scraper.py` | `generate_excel.py` | `monthly-scrape.yml` (06:00) |
| **SURE** | [certification.sure-system.org](https://certification.sure-system.org/SearchVerifications) | `scraper_sure.py` | `generate_excel_sure.py` | `monthly-scrape-sure.yml` (07:00) |
| **PEFC** | [pefc.org/find-certified-legacy](https://pefc.org/find-certified-legacy) | `scraper_pefc.py` | `generate_excel_pefc.py` | `monthly-scrape-pefc.yml` (08:00) |
| **FSC** | [FSC Certificates Public Dashboard](https://app.powerbi.com/view?r=eyJrIjoiN2U3NGMyNWEtZTAxNS00MzVhLWExNmMtOThhZjdiYjQ4MWNkIiwidCI6IjEyNGU2OWRiLWVmNjUtNDk2Yi05NmE5LTVkNTZiZWMxZDI5MSIsImMiOjl9) (Power BI) | `scraper_fsc.py` | `generate_excel_fsc.py` | `monthly-scrape-fsc.yml` (09:00) |
| **GGL** | [greengoldlabel.com/certification](https://greengoldlabel.com/certification/) (PDF) | `scraper_ggl.py` | `generate_excel_ggl.py` | `monthly-scrape-ggl.yml` (10:00, 8th) |
| **SBP** | [sbp-cert.org/certificate-holders](https://sbp-cert.org/certificate-holders/) | `scraper_sbp.py` | `generate_excel_sbp.py` | `monthly-scrape-sbp.yml` (11:00) |
| **GLOBALG.A.P** | [Supply Chain Portal](https://prod.osapiens.cloud/portal/webbundle/foodplus/field-service-os/supply-chain-portal) (osapiens) | `scraper_ggap.py` | `generate_excel_ggap.py` | `monthly-scrape-ggap.yml` (10:00) |

Each run produces `<Scheme> certificates latest.xlsx` (most recent) and a dated
`<Scheme> certificates YYYY.MM.DD.xlsx` archive. You can also trigger any scraper
manually from the **Actions** tab → *Run workflow*.

## Combined workbook

`build_combined.py` merges the six **certificate** schemes above (GLOBALG.A.P is a
producer register with a different shape, so it stays separate) into one workbook,
`All certificates latest.xlsx`, rebuilt monthly by `monthly-build-combined.yml`
(on the 8th, after all scrapes). It has:

- **Dashboard** — an interactive front page: *Select Country* → certificates by
  scheme, *Select Scheme* → top countries, plus a **Top Certification Bodies**
  chart.
- **All Certificates** — one normalised, filterable row per certificate across
  all schemes (`Scheme, Identifier, Name, Country, Type, Certification Body,
  Status, Valid From, Valid To`), with dates converted to real Excel dates so
  the whole ~200k-row set sorts and filters together.
- **Certification Bodies** — each CB, its record count, and which schemes report
  it (ISCC, SURE, GGL, SBP publish a CB; PEFC and FSC do not).
- **one sheet per scheme** — the full native columns, so no detail is lost.
- **Summary** — record counts per scheme.

The `latest` copy plus the newest dated `All certificates YYYY.MM.DD.xlsx` are
committed (each ~26 MB); the monthly rebuild prunes the previous dated copy so
only one is kept.

`email-combined.yml` is an on-demand workflow (**Actions → Run workflow**) that
emails a **link** to the latest dated workbook (not an attachment — a Gmail-sent
attachment to a Microsoft 365 inbox gets quarantined; a plain link to the public
repo does not). It reuses the same `MAIL_USERNAME` / `MAIL_PASSWORD` secrets below.

## Monthly email digest

`monthly-email-digest.yml` runs on the 8th (after every scraper has committed its
dashboard for the month) and emails all six registers as CSVs, zipped into one
attachment (~6 MB), to `maris.zamovskis@bmcertification.com`. Rather than
re-running the scrapers, `export_csvs.py` rebuilds each CSV from the committed
`latest.xlsx` dashboard's `Data` sheet, so the digest is cheap and always matches
the last committed scrape.

Sending uses Gmail SMTP via [`dawidd6/action-send-mail`](https://github.com/dawidd6/action-send-mail).
Configure two repository secrets (**Settings → Secrets and variables → Actions**):

| Secret | Value |
|---|---|
| `MAIL_USERNAME` | the Gmail address to send from |
| `MAIL_PASSWORD` | a Google [App Password](https://myaccount.google.com/apppasswords) for that account (not the normal password) |

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

## FSC

FSC's data is a **Power BI "publish to web"** report — no HTML to scrape — but the
published report exposes its data through Power BI's public `querydata` API.
`scraper_fsc.py` replays the exact semantic query behind the report's detail-table
visual and pages through it with Power BI restart tokens, so it's a **plain HTTP
scraper (no browser)**. Rows are per certificate *site*; columns:

| Column | Description |
|---|---|
| Licence Code | FSC licence code (e.g. FSC-C103661) |
| Certificate Code | Full certificate code (e.g. INT-COC-001586) |
| Certificate Type | Derived from the code (COC / FM / CW / FM/COC / …) |
| Status | Certificate status (Valid) |
| Controlled Wood | Yes / No |
| Valid From / Valid To | Certificate validity dates |
| Organization | Certificate / site holder |
| Role | Certificate holder or Site |
| Site Status, State/Province, Country | Site details |

```bash
pip install -r requirements.txt
python scraper_fsc.py
python generate_excel_fsc.py
```

## GGL

Green Gold Label has no search API or HTML register — it publishes the full
holder list as a **PDF** (exported from Excel) linked from its certification
page. `scraper_ggl.py` finds the most recent *"GGL certificate holder list"* PDF
linked on [greengoldlabel.com/certification](https://greengoldlabel.com/certification/),
downloads it, and parses the table with **pdfplumber**. Long cells (participant
name, role, CB) wrap across lines; each real row is anchored by a numeric USI in
the left column, so the parser bins every word to a column by x-position and
merges wrapped continuation lines back into their anchor row. Columns:

| Column | Description |
|---|---|
| USI | GGL unique system identifier |
| Participant name | Certificate holder |
| Country | Country |
| Participant role | e.g. *Trader*, *First collector*, *Power company* |
| Regulation | e.g. *FIT/FIP* |
| Standards | e.g. *GGLS1, GGLS4* |
| Type of biomass | e.g. *AR*, *WB*, *Cat 5* |
| Valid from / Valid till | Certificate validity dates |
| CB | Certification body |
| Status | *Valid*, *Suspended*, *Withdrawn*, *Terminated*, *Expired* |

> **Note:** unlike the other registers this list includes *all* statuses, not
> just valid ones, so the Status column is kept rather than pre-filtered.

```bash
pip install -r requirements-ggl.txt
python scraper_ggl.py
python generate_excel_ggl.py
```

## SBP

The Sustainable Biomass Program register is a WordPress *Search & Filter Pro*
directory — each holder is a server-rendered, expandable panel, 12 per page,
paginated with `?sf_paged=N`. `scraper_sbp.py` is a plain-HTTP scraper: it reads
the last page number from the first page's pagination, walks every page, and
parses each panel's detail block (a clean label/value list) plus the holder name
and country flag from the header. Columns:

| Column | Description |
|---|---|
| Certificate Number | SBP certificate code (e.g. SBP-14-06) |
| Certificate Holder | Organisation name |
| Country | Country (from the header flag) |
| Certificate Type | e.g. *Trader*, *Biomass Producer* |
| Status | *Active*, *Suspended*, *Terminated* |
| Certification Body | Issuing CB |
| Date of Issue / Date of Expiry | Certificate validity dates |
| Certificate Scope | e.g. *Includes EU RED; Includes Supply Base Evaluation* |
| Products Covered | e.g. *Wood pellets; Wood chips; …* |

> **Note:** the register lists all statuses, not just active ones, so the Status
> column is kept rather than pre-filtered.

```bash
pip install -r requirements.txt
python scraper_sbp.py
python generate_excel_sbp.py
```

## GLOBALG.A.P

Unlike the other registers, GLOBALG.A.P lists **producers** (not certificates),
via the **Supply Chain Portal** (FoodPLUS, on the osapiens platform). The portal
is a GWT single-page app whose data comes over a proprietary *binary* RPC — there
is no usable HTTP/JSON API — so `scraper_ggap.py` drives headless Chromium via
**Playwright**, like PEFC.

The portal's *Product* search **requires both a product and a country** (a
product-only search errors *"Please input Country"*), so the only way to
enumerate is to walk the grid of products × countries. The scraper therefore
scans a **configured set of countries** (default **Latvia, Estonia, Lithuania**)
across **all ~727 product options** (the full crop list, kept in
`ggap_products.json`) and collects the producer rows for each, de-duplicated by
`GGN + Country + Product`. A producer certified for several crops appears once
per crop. Columns:

| Column | Description |
|---|---|
| GGN | GLOBALG.A.P Number (the producer's unique 13-digit id) |
| Producer Name | Producer / producer-group name |
| City | City / locality |
| Country | Country (the searched country) |
| Producer Type | *Producer* or *Producer group* |
| Product | The crop the producer is certified for |

> **Note:** the public search exposes producer identity, not per-certificate
> detail (issuing CB, validity dates live on each producer's detail page), so
> those columns are omitted. The dashboard's second dimension is **Product**.

```bash
pip install -r requirements-ggap.txt
python -m playwright install chromium
python scraper_ggap.py
python generate_excel_ggap.py
```

Env vars: `GGAP_COUNTRIES="Latvia,Estonia,Lithuania"` (the scope),
`GGAP_PRODUCTS="Apple,Tomato"` (scan specific products instead of the full list),
`GGAP_MAX_PRODUCTS=20` (cap for quick tests), `GGAP_HEADFUL=1` (visible browser),
`GGAP_ENUM=1` (re-enumerate the product list into `ggap_products.json` and exit).
