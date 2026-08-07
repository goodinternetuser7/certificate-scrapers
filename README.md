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
| **RSPO** | [rspo.org/search-members](https://rspo.org/search-members/) (Salesforce) | `scraper_rspo.py` | `generate_excel_rspo.py` | `monthly-scrape-rspo.yml` (12:00) |
| **FSSC** | [fssc.com/public-register](https://www.fssc.com/public-register/) | `scraper_fssc.py` | `generate_excel_fssc.py` | `monthly-scrape-fssc.yml` (13:00) |
| **ENplus** | [enplus-pellets.eu/producer](https://enplus-pellets.eu/producer/) | `scraper_enplus.py` | `generate_excel_enplus.py` | `monthly-scrape-enplus.yml` (14:00) |

Each run produces `<Scheme> certificates latest.xlsx` (most recent) and a dated
`<Scheme> certificates YYYY.MM.DD.xlsx` archive (RSPO, a member register, uses
`RSPO members …`). You can also trigger any scraper manually from the
**Actions** tab → *Run workflow*.

The workflows are staggered an hour apart, but GitHub's cron is delayed by hours
and the long scrapes overlap, so a run can finish and find that a sibling has
already pushed to `master`. Every commit step therefore **rebases and retries**
(three times) rather than pushing once: an unguarded push is how the 2026-08-01
GLOBALG.A.P run — 74 minutes of scraping, dashboard already generated — was
thrown away to a non-fast-forward rejection.

## Combined workbook

`build_combined.py` merges the eight **certificate** schemes above (GLOBALG.A.P is a
producer register and RSPO a membership register — different shapes, so both stay
separate) into one workbook,
`All certificates latest.xlsx`, rebuilt monthly by `monthly-build-combined.yml`
(on the 8th, after all scrapes). It has:

- **Dashboard** — an interactive front page: *Select Country* → certificates by
  scheme, *Select Scheme* → top countries, plus a **Top Certification Bodies**
  chart.
- **All Certificates** — one normalised, filterable row per certificate across
  all schemes (`Scheme, Identifier, Name, Country, Type, Certification Body,
  Status, Valid From, Valid To`), with dates converted to real Excel dates so
  the whole ~240k-row set sorts and filters together.
- **Certification Bodies** — each CB, its record count, and which schemes report
  it (ISCC, SURE, GGL, SBP and ENplus publish a CB; PEFC, FSC and FSSC do not).
- **one sheet per scheme** — the full native columns, so no detail is lost.
- **Summary** — record counts per scheme.

The `latest` copy plus the newest dated `All certificates YYYY.MM.DD.xlsx` are
committed (each ~37 MB); the monthly rebuild prunes the previous dated copy so
only one is kept.

`email-combined.yml` is an on-demand workflow (**Actions → Run workflow**) that
emails a **link** to the latest dated workbook (not an attachment — a Gmail-sent
attachment to a Microsoft 365 inbox gets quarantined; a plain link to the public
repo does not). It reuses the same `MAIL_USERNAME` / `MAIL_PASSWORD` secrets below.

## Per-run email notifications

`run-notify.yml` emails one short message to
`maris.zamovskis@bmcertification.com` **every time any workflow here finishes** —
the ten scrapers, the combined-workbook build and the two email jobs — whether it
succeeded or failed, scheduled runs and manual dispatches alike:

```
Subject: [scrapers] ✅ Monthly SBP Certificate Scraper — 750 certificate holders

  Result:  Parsed 750 certificate holders (429 Active).
  Commit:  chore: monthly SBP certificate scrape 2026-08-07
  Trigger: schedule
  Run:     https://github.com/…/actions/runs/…
```

A failure instead reads `❌ …`, names the step that failed, and quotes the last
12 lines of that step's own output, so the mail usually says *why* without
opening the run.

It is a single `workflow_run` listener rather than a notify step inside each of
the thirteen workflows: one file to maintain, and it still fires when a run dies
before reaching a final step (cancelled, or the runner timing out mid-scrape) —
an in-workflow step would be skipped in exactly those cases. It is deliberately
absent from its own watch list, since a `workflow_run` trigger on itself would
recurse. The trade-off is that `workflow_run` only ever executes the
default-branch copy of the file, so it cannot be exercised from a PR branch.

Counts are read out of the finished run's log, and each job phrases its tally
differently — `Parsed N …` for most scrapers, `Wrote N rows` for GLOBALG.A.P,
`Building dashboard: N rows` for the combined build, `N CSVs, N MB` for the digest
— so the first pattern to match wins. If a log is unavailable the mail still
identifies the run, its outcome and the failing step: every probe is best-effort
by design, since an email that says only "failed, here is the link" beats
silence. Reuses the same `MAIL_USERNAME` / `MAIL_PASSWORD` secrets as the digest,
and is plain text with no attachment for the same Microsoft 365 reason.

> Without this, the only signal was GitHub's own failure email, which for a
> *scheduled* run goes to whoever last committed the workflow file — which is how
> three failed scrapes on 2026-08-01 went unnoticed until the 7th.

## Monthly email digest

`monthly-email-digest.yml` runs on the 8th (after every scraper has committed its
dashboard for the month) and emails all eight registers as CSVs, zipped into one
attachment (~11 MB), to `maris.zamovskis@bmcertification.com`. Rather than
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

> ### ⚠️ Broken since 2026-08-01 — the source was retired
>
> PEFC **deleted** the Caspio DataPage this scraper reads: it now answers
> *"DataPage does not exist"* (Caspio error 50501), and that message renders
> straight into `pefc.org/find-certified-legacy` in place of the register. The
> replacement is a new database at **`https://one.pefc.org/iframe`** (a Laravel +
> Livewire app, embedded on `pefc.org/find-certified`), but as of 2026-08-07 that
> route answers **HTTP 500** to a browser and to curl alike — the app is up
> (`/up` responds) but its register page is erroring on PEFC's side. So the
> scraper cannot yet be repointed; there is nothing scrapeable at either address.
>
> `scraper_pefc.py` now **pre-flights the DataPage and exits in under a second**
> with that explanation, instead of burning seven minutes on six identical
> 60-second Playwright timeouts and reporting "union still short after 6 passes",
> which read like scraper flakiness. The monthly workflow is left on its schedule
> so the failure stays visible. **`PEFC certificates latest.xlsx` is therefore
> frozen at the 2026-07-01 scrape**, and the combined workbook keeps merging that
> snapshot.
>
> Everything below describes the retired Caspio source and is kept for whoever
> repoints the scraper.

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

> **Blank pages:** on 2026-08-01 the register answered `200 OK` with a page
> carrying no holder panels at all; the scraper read that as "no results" and
> failed the run in 18 seconds. A re-run on unchanged code scraped all 750
> holders, so it was a blip, not a layout change. Because paging past the end
> re-serves earlier holders rather than an empty page, a holder-less page is
> *always* anomalous — so it is now retried four times (and what came back is
> logged: status, size, title) before the run fails. That also closes a latent
> hole: a blank page in the *middle* of the walk used to end the loop quietly
> and commit a truncated register. As a second net, the run fails if it stops
> more than one page short of the highest page the pager advertised.

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
scans a **configured set of countries** (default **Latvia, Estonia, Lithuania,
Finland**)
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

> **Finland** was added to the scope on 2026-08-07. It is a much denser market
> than the Baltics: a probe over eight Nordic crops returned **27 producer-rows
> for Finland against 4 for Latvia**, so it materially grows the dataset (the
> Baltics-only scrape was 99 rows). Cost is one more pass over the 727 products
> — roughly **+25 min**, taking a full run from ~1h12m to ~1h35m, well inside the
> workflow's 330-minute timeout. The country name must match the portal's own
> dropdown option exactly, which "Finland" does.

```bash
pip install -r requirements-ggap.txt
python -m playwright install chromium
python scraper_ggap.py
python generate_excel_ggap.py
```

Env vars: `GGAP_COUNTRIES="Latvia,Estonia,Lithuania,Finland"` (the scope),
`GGAP_PRODUCTS="Apple,Tomato"` (scan specific products instead of the full list),
`GGAP_MAX_PRODUCTS=20` (cap for quick tests), `GGAP_HEADFUL=1` (visible browser),
`GGAP_ENUM=1` (re-enumerate the product list into `ggap_products.json` and exit).

## RSPO

RSPO publishes **members**, not certificates. `rspo.org/search-members/` is only a
WordPress shell around an iframe to a Salesforce Visualforce page
(`rspo.my.salesforce-sites.com/membership/AT_SearchMember_VFPage`), which boots a
Lightning Out component — so there is no HTML register to parse. The component's
data comes from a **guest-accessible Apex action**, so `scraper_rspo.py` is a
**plain HTTP scraper (no browser)**: it bootstraps the Aura framework context from
the Lightning Out app descriptor (the `fwuid` rotates on every Salesforce deploy,
so it must be read per run), then POSTs one `getApplicationsByFilter` action with
an empty filter and `queryLimit = 0` ("no limit"). The whole register (~6.4k
members) comes back in a single ~3 MB response in a couple of seconds — the
cheapest scraper here.

The action also returns its own `RecordCount`, so the pull is self-checking: a
short response is retried, then falls back to unioning per-membership-category
pulls. Columns:

| Column | Description |
|---|---|
| Membership Number | RSPO membership number (e.g. 2-0516-14-000-00) |
| Member Name | Member organisation |
| Country | Country / territory |
| Membership Category | *Ordinary*, *Associate* or *Affiliate* |
| Sector | e.g. *Oil Palm Growers*, *Retailers*, *Supply Chain Associate* |
| Status | *Active* or *Suspended* |
| Last Update | Date the membership record was last updated |
| Group Members | Number of group members under the membership |
| Group Member Names | Their names, `;`-joined (the register lists no other detail) |
| Profile URL | Public member profile on rspo.org (blank if the member has none) |

> **Note:** the register lists both Active and Suspended members, so the Status
> column is kept rather than pre-filtered. Members are not certificates — there
> is no CB, certificate number or validity window to scrape — so RSPO stays out
> of the combined workbook, like GLOBALG.A.P.

```bash
pip install -r requirements.txt
python scraper_rspo.py
python generate_excel_rspo.py
```

## FSSC

The FSSC public register (FSSC 22000 food safety + FSSC 24000 social management)
is a Vue app on a WordPress site, and its table is rendered from a
guest-accessible **admin-ajax** action (`certificate_getCertificates`) — so
`scraper_fssc.py` is a **plain HTTP scraper (no browser)**. Three quirks shape it,
all measured against the live site: Cloudflare answers **403** to a non-browser
User-Agent; `limit` is capped **server-side at 15** (asking for more still
returns 15), so the ~42k organizations take ~2.8k paged requests; and Cloudflare
rate-limits the endpoint at roughly **75 requests/minute** per IP (429 with
`Retry-After: ~125`). Requests are therefore paced at one per 1.1s (≈55/min,
measured as sustainable), so a full run takes **~55 minutes**. A 429 isn't fatal:
the scraper honours `Retry-After` and permanently slows itself down before
retrying that offset.

The register is offset-paged over a live, alphabetically sorted database that is
updated daily, so a row inserted or removed mid-run shifts every later offset by
one and can duplicate or skip a record at the seam. Records are therefore
de-duplicated by COID and the final count is checked against the API's own
`total`; a shortfall beyond 0.5% fails the run rather than committing a partial
register. Columns:

| Column | Description |
|---|---|
| COID | FSSC certified-organization ID (e.g. UZB-1-6756-503556) |
| Organization | Certified organization |
| Country / City | Site address |
| Scheme | *FSSC 22000* or *FSSC 24000* |
| Status | *Valid* or *Suspended* |
| Food Chain Category | e.g. *CIV : Processing of ambient stable products* |
| Product Types | e.g. *Packaging*, `;`-joined if several |
| Scope Statement | The certified scope, as published |
| Initial Certification | Date of the first certification |
| Issued / Valid Until | Current certificate validity dates |
| Last Status Decision | Date of the last status decision (*"suspended since"*) |
| GFSI Recognized | Whether the certification is GFSI-recognized |

> **Note:** the register lists both Valid and Suspended organizations, so the
> Status column is kept rather than pre-filtered. It publishes no issuing
> Certification Body (that is a *Public Register Plus* feature), so — like PEFC
> and FSC — that column is omitted and the dashboard's second dimension is the
> food chain **Category**. A quarter of organizations hold several categories;
> the dashboard keys on the *first* one, so its selector stays the register's own
> list of 16 categories rather than ~200 combination strings, while the Data
> sheet keeps the full list per row. A few published scope statements and city
> names contain control characters that Excel rejects — the scraper strips those.

```bash
pip install -r requirements.txt
python scraper_fssc.py
python generate_excel_fssc.py
```

Env vars for local runs: `FSSC_MAX_PAGES=20` (cap pages for a quick test),
`FSSC_DELAY=1.1` (seconds between requests), `FSSC_PAGE_SIZE=15`.

## ENplus

The ENplus® producer register is a WordPress page whose table is rendered by the
site's own `enplus_category_list` plugin over **admin-ajax**, so `scraper_enplus.py`
is a **plain HTTP scraper (no browser, no credentials)** and a full run takes
**under two minutes**. Two actions are involved:

- `handle_category_search_request` returns the results table as an HTML fragment.
  Its `data` parameter is the search form's own jQuery serialisation, in which
  the hidden `category` field picks the register (17 = Producer) and the
  `certificate_status[]` boxes are 1 = Active, 2 = Suspended, 3 = Terminated.
  **All three are ticked by default**, so — unlike the other scrapers here — this
  register is not a snapshot of valid certificates but the full history, ~42% of
  it terminated. Those ids are read off the live page each run rather than
  hard-coded, since a plugin rebuild could renumber them.
- `handle_cert_request` returns one company's popup. It is the only place the
  **certification body**, quality classes, certified activities, legal address,
  certified sites and approved bag designs appear, so every company is fetched
  (six at a time) and merged into its row. The popup repeats the ENplus ID, which
  doubles as a check that the merge lined up: any mismatch fails the run.

The register reports page counts rather than a total, so the completeness check is
that the row count falls inside the range those pages imply; the pull is also
unioned by company id, so a row shifting between pages cannot duplicate a company.

| Column | Description |
|---|---|
| ENplus ID | ENplus® company id, country-prefixed (e.g. `AT 001`) |
| Producer | Company name, as published |
| Status | *Active*, *Suspended* or *Terminated* |
| Status Since | Date of that status — only published when not Active |
| Country / City | From the register's table |
| Certification Body | The certifying body (ENplus publishes this; most schemes here do not) |
| Quality Classes | e.g. *ENplus® A1 - 6mm*, `;`-joined if several |
| Certified Activities | Production, bagging, large-scale delivery, storage … |
| Legal Address | From the company popup |
| Certified Sites / Certified Site Names | Site count, and each site with its own status — a site can be suspended while its company is active |
| Bag Designs | Number of approved bag designs |
| Website / Company ID | Official site; internal id used to re-fetch the popup |

> **Note:** the register carries **no certificate numbers and no validity dates**
> — the "Since" date of a suspension or termination is the only date published.
> Older terminated entries are also thinly populated (about 4% have no
> certification body and 18% no certified activities, almost all of them
> terminated); those gaps are in the source and are left blank rather than
> guessed. Four companies have no country in the table at all.

```bash
pip install -r requirements.txt
python scraper_enplus.py
python generate_excel_enplus.py
```

Env vars for local runs: `ENPLUS_MAX_PAGES=1` (cap pages for a quick test),
`ENPLUS_WORKERS=6` (concurrent popup fetches), `ENPLUS_SKIP_DETAILS=1` (list only,
no popups).
