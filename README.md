# ISCC Certificate Scraper

Fetches all **active (valid)** ISCC certificates from [iscc-system.org](https://iscc-system.org/certification/all-certificates/) once a month and stores them as a CSV.

## Columns

| Column | Description |
|---|---|
| Client Name | Organisation holding the certificate |
| Scope | Certification scope (e.g. Processing Unit, Trader) |
| Issuing CB | Certification body that issued the certificate |
| Expiry Date | Certificate expiry date (DD.MM.YYYY) |

## Output files

- `certificates_latest.csv` — always the most recent run
- `certificates_YYYY-MM-DD.csv` — dated archive of each run

## Schedule

Runs automatically at **06:00 UTC on the 1st of every month** via GitHub Actions.  
You can also trigger it manually from the **Actions** tab → *Monthly ISCC Certificate Scraper* → *Run workflow*.

## Running locally

```bash
pip install -r requirements.txt
python scraper.py
```
