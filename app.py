import io
import os
import re
import tempfile
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Daily Report Automation", page_icon="📊", layout="wide")

# Optional password protection. Configure APP_PASSWORD in Streamlit Secrets.
APP_PASSWORD = st.secrets.get("APP_PASSWORD", os.getenv("APP_PASSWORD", ""))

def norm(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def key(value):
    return re.sub(r"\s+", " ", norm(value)).lower()

def task_type(task_no):
    value = norm(task_no).upper()
    if "TKT" in value:
        return "TKT"
    if "POI" in value:
        return "POI"
    return ""

def classify(task_kind, service_area, root_cause):
    kind = norm(task_kind).upper()
    service = norm(service_area).upper()
    root = norm(root_cause).upper()
    if kind == "TKT":
        if "FR-SLA" in service or "BIZ" in service:
            return "FR-SLA / Biz TKT"
        if "INSTALLATION TYPE CHANGE" in root:
            return "FR-SLA / Biz TKT"
        if "FT-SBS" in root:
            return "FT-SBS"
        return "Other TKT"
    if kind == "POI":
        return "Install Express" if "EXPRESS" in root else "Install POI"
    return ""

def process_sheet(df, sheet_name):
    df = df.dropna(how="all").copy()
    # Keep the user's original columns; add calculated fields without replacing raw data.
    cols = {key(c): c for c in df.columns if c is not None}
    task_col = cols.get("task no")
    service_col = cols.get("service area")
    root_col = cols.get("root cause")
    status_col = cols.get("task status")
    repay_col = cols.get("repay status")
    eng_col = cols.get("lan engineer name")
    car_col = cols.get("car id")
    labor_col = cols.get("contractor / office labour id")

    if task_col:
        kinds = df[task_col].map(task_type)
        df["TKT / POI"] = kinds
        df["Classification"] = [
            classify(k, s, r)
            for k, s, r in zip(
                kinds,
                df[service_col] if service_col else [""] * len(df),
                df[root_col] if root_col else [""] * len(df),
            )
        ]

    if car_col and eng_col:
        eligible = df[[car_col, eng_col]].copy()
        eligible.columns = ["car", "engineer"]
        eligible["car"] = eligible["car"].map(norm)
        eligible["engineer"] = eligible["engineer"].map(norm)
        if repay_col:
            repay_values = df[repay_col].map(norm).str.lower()
            excluded = repay_values.str.contains(r"engr\s*leave|route\s*cancel", regex=True, na=False)
            eligible = eligible.loc[~excluded]
        eligible = eligible[(eligible["car"] != "") & (eligible["engineer"] != "")]
        car_people = eligible.groupby("car")["engineer"].apply(lambda values: len(set(values))).to_dict()
        df["RC"] = df[car_col].map(lambda car: (
            "" if not norm(car) else
            "Car Share Plus" if car_people.get(norm(car), 0) >= 3 else
            "Car Share" if car_people.get(norm(car), 0) == 2 else
            "One Car" if car_people.get(norm(car), 0) == 1 else ""
        ))

    # Preserve Task Status as text; do not alter Schedule Date.
    if status_col:
        df[status_col] = df[status_col].map(norm)

    metrics = summarize_sheet(df, sheet_name)
    return df, metrics

RESOLVED_STATUSES = {
    "Resolved", "Resolved (No KPI)", "Resolved (AP Up)", "Resolved (ONU Up)"
}
VISITED_TKT_STATUSES = {
    "Engr Visited Pending", "Handover to FTTx", "Handover to QC-Wireless",
    "Handover to S-Combine", "Handover to Sales", "Resolved",
    "Resolved (No KPI)", "Resolved (AP Up)", "Resolved (ONU Up)",
    "Unresolved(Cabling Done)"
}

def summarize_sheet(df, sheet_name):
    cols = {key(c): c for c in df.columns if c is not None}
    task_col = cols.get("task no")
    status_col = cols.get("task status")
    eng_col = cols.get("lan engineer name")
    car_col = cols.get("car id")
    labor_col = cols.get("contractor / office labour id")
    repay_col = cols.get("repay status")
    kinds = df[task_col].map(task_type) if task_col else pd.Series([""] * len(df), index=df.index)
    statuses = df[status_col].map(norm) if status_col else pd.Series([""] * len(df), index=df.index)
    task_present = df[task_col].map(norm).ne("") if task_col else pd.Series([False] * len(df), index=df.index)
    resolved = int((task_present & statuses.isin(RESOLVED_STATUSES)).sum())
    visited = int((task_present & kinds.eq("TKT") & statuses.isin(VISITED_TKT_STATUSES)).sum())
    engineers = int(df[eng_col].map(norm).replace("", pd.NA).nunique()) if eng_col else 0
    cars = int(df[car_col].map(norm).replace("", pd.NA).nunique()) if car_col else 0
    if labor_col:
        labor_values = df[labor_col].map(norm)
        if repay_col:
            excluded = df[repay_col].map(norm).str.lower().str.contains(r"engr\s*leave", regex=True, na=False)
            labor_values = labor_values.loc[~excluded]
        labor_values = labor_values[~labor_values.str.lower().isin({"", "no cabler", "no labor"})]
        labor = int(labor_values.nunique())
    else:
        labor = 0
    return {
        "Raw Sheet": sheet_name,
        "Rows with Task No": int(task_present.sum()),
        "TKT": int((task_present & kinds.eq("TKT")).sum()),
        "POI": int((task_present & kinds.eq("POI")).sum()),
        "Resolved (KPI statuses)": resolved,
        "Visited TKT": visited,
        "Total Team (unique LAN Engineer Name)": engineers,
        "Car (unique Car ID)": cars,
        "Labor (excluding No Cabler/No Labor)": labor,
    }

def excel_date_label(raw_frames):
    dates = []
    for df in raw_frames.values():
        col = next((c for c in df.columns if key(c) == "schedule date"), None)
        if col:
            for value in df[col].dropna().tolist():
                if isinstance(value, (datetime, date, pd.Timestamp)):
                    dates.append(pd.Timestamp(value).date())
                else:
                    parsed = pd.to_datetime(value, errors="coerce")
                    if not pd.isna(parsed):
                        dates.append(parsed.date())
    if not dates:
        return date.today().isoformat()
    # Use the most common date, avoiding header/crew rows without dates.
    return pd.Series(dates).value_counts().index[0].isoformat()

def style_sheet(ws):
    navy = "18212F"
    white = "FFFFFF"
    thin = Side(style="thin", color="D9DEE6")
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(color=white, bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells[:250])
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 36)

def make_output_report(report_bytes, metrics_df, report_date):
    source = io.BytesIO(report_bytes)
    wb = load_workbook(source)
    sheet_name = "Auto Generated Summary"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.append(["DAILY REPORT AUTOMATION — GENERATED SUMMARY"])
    ws.append(["Raw data date", report_date])
    ws.append(["Note", "This summary is calculated from the uploaded Raw Excel. Existing report sheets are preserved unchanged until their exact cell mapping is confirmed."])
    ws.append([])
    ws.append(list(metrics_df.columns))
    for row in metrics_df.itertuples(index=False, name=None):
        ws.append(list(row))
    # Style the title and table.
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, len(metrics_df.columns)))
    ws["A1"].fill = PatternFill("solid", fgColor="18212F")
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=14)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 28
    for cell in ws[5]:
        cell.fill = PatternFill("solid", fgColor="334155")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in ws.iter_rows(min_row=6):
        for cell in row:
            cell.border = Border(bottom=Side(style="hair", color="D9DEE6"))
            if cell.row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F1F3F6")
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:{get_column_letter(len(metrics_df.columns))}{ws.max_row}"
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 13), 42)
    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue()

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if APP_PASSWORD:
    if not st.session_state.authenticated:
        st.title("🔐 Daily Report Automation")
        pw = st.text_input("Password", type="password")
        if st.button("Login", type="primary"):
            if pw == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Password မှားနေပါတယ်။")
        st.stop()
else:
    st.info("Password protection ဖွင့်ရန် Streamlit Secrets ထဲတွင် APP_PASSWORD ထည့်ပါ။")

st.markdown("# 📊 DAILY REPORT AUTOMATION")
st.caption("Automated Excel Reporting System")
st.markdown("**Raw Excel → Process → Report Excel**")
st.warning(
    "ဒီ App ဗားရှင်းက Raw Data ကိုတွက်ချက်ပြီး Report နမူနာဖိုင်ထဲတွင် "
    "'Auto Generated Summary' Sheet အသစ်ထည့်ပေးသည်။ မူရင်း Report Sheet ၉ ခုကို "
    "မှားယွင်းသော cell များသို့ အလိုအလျောက်မရေးသေးပါ။ အတိအကျ cell mapping အတည်ပြုပြီးမှ "
    "ထို Sheet များကို auto-fill လုပ်မည်။"
)

left, right = st.columns(2)
with left:
    raw_file = st.file_uploader("၁။ Raw Excel ဖိုင်တင်ပါ", type=["xlsx"], key="raw")
with right:
    report_file = st.file_uploader("၂။ Report 07-Oct နမူနာ Excel တင်ပါ", type=["xlsx"], key="report")

if raw_file and report_file:
    if st.button("🚀 GENERATE REPORT", type="primary", use_container_width=True):
        try:
            raw_bytes = raw_file.getvalue()
            raw_book = pd.ExcelFile(io.BytesIO(raw_bytes))
            required_sheets = ["R1 OC", "R6", "R6 OC", "R1", "R1S", "R2"]
            raw_frames = {}
            processed_frames = {}
            metrics = []
            aliases = {"R1 OC": ["R1 OC", "R1OC"], "R6 OC": ["R6 OC", "R6OC"]}
            for sheet in required_sheets:
                actual = sheet
                if actual not in raw_book.sheet_names and sheet in aliases:
                    actual = next((name for name in aliases[sheet] if name in raw_book.sheet_names), None)
                if actual and actual in raw_book.sheet_names:
                    frame = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=actual)
                    raw_frames[sheet] = frame
                    processed, summary = process_sheet(frame, sheet)
                    processed_frames[sheet] = processed
                    metrics.append(summary)
            if not metrics:
                st.error("Raw Excel ထဲမှာ လိုအပ်တဲ့ Sheet တွေ မတွေ့ပါ။")
                st.stop()

            report_date = excel_date_label(raw_frames)
            metrics_df = pd.DataFrame(metrics)
            output_report = make_output_report(report_file.getvalue(), metrics_df, report_date)

            # Also provide a processed-data workbook as a separate download.
            processed_buffer = io.BytesIO()
            with pd.ExcelWriter(processed_buffer, engine="openpyxl") as writer:
                for sheet, frame in processed_frames.items():
                    frame.to_excel(writer, sheet_name=sheet[:31], index=False)
            st.success(f"Report ဖိုင် ပြင်ဆင်ပြီးပါပြီ။ Raw data date: {report_date}")
            st.dataframe(metrics_df, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Download Report Excel",
                data=output_report,
                file_name=f"Daily_Report_{report_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.download_button(
                "⬇️ Download Processed Raw Data Excel",
                data=processed_buffer.getvalue(),
                file_name=f"Processed_Raw_Data_{report_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Report ထုတ်ရာတွင် အမှားဖြစ်နေပါတယ်: {exc}")
            st.exception(exc)
else:
    st.info("Excel ဖိုင်နှစ်ခုလုံးတင်ပြီးမှ Generate Report ကိုနှိပ်ပါ။")
