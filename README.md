# Daily Report Automation — New App

This is a separate Streamlit app. It does not replace or modify the existing deployed app.

## Files
- `app.py` — new two-upload Streamlit app
- `requirements.txt` — Python dependencies
- `.streamlit/secrets.toml.example` — example password configuration

## What this first version does
1. Upload a Raw Excel workbook with the six expected raw sheets.
2. Upload the filled `Report 07-Oct` sample workbook.
3. Calculates per-raw-sheet TKT, POI, KPI Resolved, Visited TKT, unique team, car, and labor counts.
4. Adds an `Auto Generated Summary` sheet to a copy of the report sample and preserves the existing report sheets.
5. Offers a second download with enriched raw sheets and the TKT/POI, Classification, and RC columns.

## Important limitation
This version does NOT write calculated values into the existing nine report sheets because the exact mapping between raw metrics and report cells has not yet been confirmed. It deliberately preserves those sheets unchanged instead of guessing and potentially corrupting the report. Once cell-by-cell mapping is confirmed, `make_output_report()` can be extended to fill the intended cells.

## Deploy on Streamlit Community Cloud
1. Create a new GitHub repository, e.g. `daily-report-automation-new`.
2. Upload `app.py`, `requirements.txt`, and `.streamlit/secrets.toml.example` (the example is not a real secret).
3. In Streamlit Community Cloud, create a new app from this new repository and choose `app.py`.
4. In App Settings → Secrets, set:
   `APP_PASSWORD = "your-own-password"`
5. Deploy.

Do not change the existing repository or deployed app.
