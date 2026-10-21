# Jumia KE – Daily Prohibited Items Monitor

Checks jumia.co.ke every morning against the Prohibited / Restricted Brands Google Sheet and emails an Excel report so listings can be followed up for delisting.

## How it works
1. Loads rules live from the Google Sheet (restricted brands, pharma OTC/Rx, named skin-lightening products) plus `PROHIBITED_KEYWORDS` in `monitor.py`.
2. Searches Jumia for each term (newest first) and matches listings.
3. Classifies hits using the Allowed Vendors column:
   - **DELIST** – no approved vendor, or seller not on the allowed list
   - **REVIEW** – "Central Restricted Brands" or seller could not be identified
   - skipped – seller is an authorised vendor
4. Compares with the previous run: **New today**, **Days live**, **Resolved since yesterday**.
5. Writes `reports/jumia_prohibited_YYYY-MM-DD.xlsx` and emails it.

## Setup
1. Share the Google Sheet as **Anyone with the link – Viewer**.
2. Push this folder to a **private** GitHub repo.
3. Repo > Settings > Secrets and variables > Actions, add:
   - `SMTP_USER` – Gmail address that sends the report
   - `SMTP_PASS` – Gmail **app password** (Google Account > Security > 2-Step Verification > App passwords)
   - `REPORT_TO` – recipient(s), comma separated
4. Actions tab > "Jumia prohibited items check" > **Run workflow** to test.
5. It then runs daily at 03:30 UTC (06:30 EAT). Edit the cron in `.github/workflows/daily.yml` to change.

## Run locally
```
pip install -r requirements.txt
python monitor.py                      # saves report only if SMTP vars not set
SHEET_XLSX=rules.xlsx python monitor.py   # use a downloaded copy of the sheet
```

## Must verify on first run
- CSS selectors in `search()` (product cards) and `seller_of()` (seller name) were written without access to the live site.
- `sort=newest` URL parameter.
- Jumia may rate-limit/block automated traffic. Tune `DELAY` and `MAX_PAGES` (env vars) or ask for an official data export.

## Tuning false positives
Add generic words to `NOISE` in `monitor.py`; for those (and any term of 3 characters or fewer) a listing is only flagged when its Brand field matches exactly.
