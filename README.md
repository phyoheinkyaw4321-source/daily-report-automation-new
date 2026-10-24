# Daily Report Automation

A separate Streamlit app for calculating daily metrics from an Auto_Pivot workbook while preserving the uploaded Report Excel.

## Upload
1. Auto_Pivot Excel containing `R1`, `R1S`, `R2`, and `R6` data sheets.
2. Filled Report Excel sample to use as the output base.

## Current rules
- `Visited TKT`: TKT rows with one of the 10 specified visited statuses.
- `Resolved`: only `Resolved`, `Resolved (No KPI)`, `Resolved (AP Up)`, `Resolved (ONU Up)`.
- `Remote Resolved (Auto)`: count `Resolved (Auto)` separately.
- `Total Team`: unique `LAN Engineer Name`.
- `Car`: unique `Car ID`.
- `Labor`: unique labor IDs excluding blank, `No Cabler`, and `No Labor`; exclude `Engr Leave` rows.
- Included areas: `R1 (wo R1S)`, `R1S`, `R2 (wo R2S)`, `R6`.
- `R1 OCV` and `R6 OCV` are deferred for now.

## Important
This version creates a professionally formatted `Automation Summary` sheet and preserves all existing Report sheets. It does not yet write into the Report's existing cells because the exact destination cells/date-column behavior has not been verified. It warns the user before download so data is not written to the wrong cells.

## Deploy
Upload `app.py`, `requirements.txt`, and `README.md` to your new GitHub repository, then create/redeploy the Streamlit app from that repository. For password protection, set `APP_PASSWORD` in Streamlit Secrets; do not put your real password in GitHub.
