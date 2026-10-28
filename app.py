import io
import os
import re
from datetime import date, datetime

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Daily Report Automation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# -------------------- Security --------------------
APP_PASSWORD = st.secrets.get("APP_PASSWORD", os.getenv("APP_PASSWORD", ""))

# -------------------- Calculation rules --------------------
RESOLVED_STATUSES = {
    "resolved",
    "resolved (no kpi)",
    "resolved (ap up)",
    "resolved (onu up)",
}
VISITED_TKT_STATUSES = {
    "engr visited pending",
    "handover to fttx",
    "handover to qc-wireless",
    "handover to s-combine",
    "handover to sales",
    "resolved",
    "resolved (no kpi)",
    "resolved (ap up)",
    "resolved (onu up)",
    "unresolved(cabling done)",
}
REGION_SHEETS = [
    ("R1 V", "R1 (wo R1S)"),
    ("R1S V", "R1S"),
    ("R2 V", "R2 (wo R2S)"),
    ("R6 V", "R6"),
]
# R1 OCV and R6 OCV intentionally left out for now.


def clean(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def norm_key(value):
    return re.sub(r"\s+", " ", clean(value)).casefold()


def find_column(df, name):
    wanted = norm_key(name)
    for col in df.columns:
        if norm_key(col) == wanted:
            return col
    return None


def date_label(frames):
    values = []
    for df in frames.values():
        col = find_column(df, "Schedule Date")
        if col is None:
            continue
        for val in df[col].dropna().tolist():
            try:
                dt = pd.to_datetime(val, errors="coerce")
                if not pd.isna(dt):
                    values.append(dt.date())
            except Exception:
                pass
    if not values:
        return date.today().isoformat()
    return pd.Series(values).value_counts().index[0].isoformat()


def summarize(df, area):
    status_col = find_column(df, "Task Status")
    task_col = find_column(df, "Task No")
    eng_col = find_column(df, "LAN Engineer Name")
    car_col = find_column(df, "Car ID")
    labor_col = find_column(df, "Contractor / Office Labour ID")

    if status_col:
        statuses = df[status_col].map(norm_key)
    else:
        statuses = pd.Series([""] * len(df), index=df.index)
    if task_col:
        tasks = df[task_col].map(clean)
        task_present = tasks.ne("")
        tkt_mask = tasks.str.contains("TKT", case=False, na=False)
        poi_mask = tasks.str.contains("POI", case=False, na=False)
    else:
        task_present = pd.Series([False] * len(df), index=df.index)
        tkt_mask = task_present.copy()
        poi_mask = task_present.copy()

    visited = int((task_present & tkt_mask & statuses.isin(VISITED_TKT_STATUSES)).sum())
    resolved = int((task_present & statuses.isin(RESOLVED_STATUSES)).sum())
    remote_resolved = int((task_present & statuses.eq("resolved (auto)")).sum())

    teams = int(df[eng_col].map(clean).replace("", pd.NA).nunique()) if eng_col else 0
    cars = int(df[car_col].map(clean).replace("", pd.NA).nunique()) if car_col else 0

    if labor_col:
        labor_values = df[labor_col].map(clean)
        labor_values = labor_values[~labor_values.map(norm_key).isin({"", "no cabler", "no labor"})]
        # Engr Leave rows are excluded when the Repay Status field identifies them.
        repay_col = find_column(df, "Repay Status")
        if repay_col:
            leave = df[repay_col].map(norm_key).str.contains(r"engr\s*leave", regex=True, na=False)
            labor_values = labor_values.loc[~leave]
        labor = int(labor_values.nunique())
    else:
        labor = 0

    return {
        "Report Area": area,
        "Visited TKT": visited,
        "Resolved": resolved,
        "Remote Resolved (Auto)": remote_resolved,
        "Total Team": teams,
        "Car": cars,
        "Labor": labor,
        "TKT": int((task_present & tkt_mask).sum()),
        "POI": int((task_present & poi_mask).sum()),
    }


def style_summary(ws, table_header_row, last_col):
    navy = "172033"
    slate = "334155"
    white = "FFFFFF"
    pale = "F3F5F8"
    border_color = "D9E0E8"
    thin = Side(style="thin", color=border_color)

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f"A{table_header_row + 1}"
    ws.auto_filter.ref = f"A{table_header_row}:{get_column_letter(last_col)}{ws.max_row}"

    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(name="Aptos Display", size=15, bold=True, color=white)
        cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30

    for cell in ws[table_header_row]:
        cell.fill = PatternFill("solid", fgColor=slate)
        cell.font = Font(name="Aptos", bold=True, color=white)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    ws.row_dimensions[table_header_row].height = 32

    for row in ws.iter_rows(min_row=table_header_row + 1):
        for cell in row:
            cell.font = Font(name="Aptos", size=10, color="202938")
            cell.border = Border(bottom=Side(style="hair", color=border_color))
            if cell.row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=pale)
            if cell.column > 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
    for col_cells in ws.columns:
        letter = get_column_letter(col_cells[0].column)
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells[:250])
        ws.column_dimensions[letter].width = min(max(max_len + 3, 14), 34)


def build_report(report_bytes, metrics, report_date):
    wb = load_workbook(io.BytesIO(report_bytes))
    sheet_name = "Automation Summary"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.append(["DAILY REPORT AUTOMATION | SUMMARY"])
    ws.append(["Data date", report_date])
    ws.append(["Included areas", "R1 (wo R1S), R1S, R2 (wo R2S), R6"])
    ws.append(["Calculation note", "R1 OCV and R6 OCV are intentionally excluded for now. Existing report sheets are preserved."])
    ws.append([])
    ws.append(list(metrics.columns))
    for row in metrics.itertuples(index=False, name=None):
        ws.append(list(row))
    style_summary(ws, 6, len(metrics.columns))
    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue()


# -------------------- Login --------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if APP_PASSWORD:
    if not st.session_state.authenticated:
        st.markdown(
            """
            <div class="login-wrap">
              <div class="login-card">
                <div class="eyebrow">SECURE WORKSPACE</div>
                <h1>Daily Report Automation</h1>
                <p>Sign in to generate the daily reporting workbook.</p>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        password = st.text_input("Access password", type="password", placeholder="Enter your password")
        if st.button("Sign in", type="primary", use_container_width=True):
            if password == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Password မမှန်ပါ။ ထပ်စစ်ပြီး ဝင်ပါ။")
        st.stop()

# -------------------- Professional UI --------------------
st.markdown(
    """
    <style>
      :root { color-scheme: light; }
      .stApp { background: #F1F4F8; color: #202938; }
      [data-testid="stHeader"] { background: rgba(241,244,248,0.98); }
      .block-container { max-width: 1240px; padding-top: 1.25rem; padding-bottom: 3rem; }
      /* High-contrast corporate header */
      .hero {
        background: linear-gradient(120deg, #101827 0%, #26364D 100%);
        color: #FFFFFF; padding: 30px 34px; border-radius: 18px;
        margin-bottom: 22px; box-shadow: 0 10px 28px rgba(15,23,42,.16);
        border: 1px solid #26364D;
      }
      .hero .eyebrow { color: #E2E8F0; font-size: 12px; letter-spacing: 2px; font-weight: 900; }
      .hero h1 { color: #FFFFFF !important; margin: 9px 0 8px 0; font-size: clamp(25px, 3vw, 34px); line-height: 1.2; font-weight: 900; }
      .hero p { color: #F8FAFC !important; margin: 0; font-size: 16px; font-weight: 600; }
      .section-label { font-size: 14px; font-weight: 900; letter-spacing: 1.2px; color: #26364D; margin: 24px 0 12px 0; }
      /* Force readable text throughout the page */
      .stApp, .stApp p, .stApp label, .stApp span, .stApp div,
      .stApp [data-testid="stMarkdownContainer"] p { color: #202938; }
      .stApp [data-testid="stMarkdownContainer"] strong { color: #101827 !important; font-weight: 900 !important; }
      .stApp [data-testid="stFileUploader"] { background: #FFFFFF; border: 1px solid #C7D2E0; border-radius: 14px; padding: 16px; }
      .stApp [data-testid="stFileUploader"] label,
      .stApp [data-testid="stFileUploader"] label p { color: #182334 !important; font-size: 15px !important; font-weight: 800 !important; }
      .stApp [data-testid="stFileUploaderDropzone"] { background: #F8FAFC !important; border: 2px dashed #A8B7C9 !important; border-radius: 10px !important; }
      .stApp [data-testid="stFileUploaderDropzone"] * { color: #26364D !important; font-weight: 700 !important; }
      .stApp [data-testid="stFileUploaderDropzone"] button { background: #FFFFFF !important; color: #172033 !important; border: 1px solid #64748B !important; font-weight: 800 !important; }
      .stApp [data-testid="stFileUploaderDropzone"] small { color: #475569 !important; }
      div.stButton > button, div.stDownloadButton > button {
        border-radius: 11px; min-height: 50px; font-weight: 900 !important;
        font-size: 16px !important; border: 1px solid #16263D;
      }
      div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {
        background: #17263D !important; color: #FFFFFF !important;
      }
      div.stButton > button[kind="primary"] p, div.stDownloadButton > button p { color: #FFFFFF !important; font-weight: 900 !important; }
      div.stButton > button:disabled { background: #D7E0EA !important; color: #526174 !important; border: 1px solid #BCC9D8 !important; opacity: 1 !important; }
      div.stButton > button:disabled p { color: #526174 !important; font-weight: 900 !important; }
      [data-testid="stAlert"] { border-radius: 12px; font-weight: 700; }
      [data-testid="stAlert"] p { color: #1E293B !important; font-size: 15px !important; font-weight: 700 !important; }
      div[data-testid="stMetric"] { background: #FFFFFF; border: 1px solid #D5DEE8; border-radius: 12px; padding: 16px; }
      div[data-testid="stMetricLabel"] * { color: #475569 !important; font-weight: 800 !important; }
      div[data-testid="stMetricValue"] * { color: #101827 !important; font-weight: 900 !important; }
      .footnote { color: #475569 !important; font-size: 13px; font-weight: 600; }
      @media (max-width: 700px) {
        .hero { padding: 23px 20px; }
        .hero h1 { font-size: 25px; }
        .hero p { font-size: 14px; }
      }
    </style>
    <div class="hero">
      <div class="eyebrow">REPORTING WORKSPACE</div>
      <h1>📊 DAILY REPORT AUTOMATION</h1>
      <p>Automated Excel Reporting System · Upload → Calculate → Download</p>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    "<div class='section-label'>01 / UPLOAD FILES</div>",
    unsafe_allow_html=True,
)
left, right = st.columns(2, gap="large")
with left:
    st.markdown("### **01. Auto_Pivot Excel**")
    pivot_file = st.file_uploader(
        "📂 Choose Auto_Pivot Excel file",
        type=["xlsx"],
        key="pivot_file",
        help="Auto_Pivot workbook containing R1, R1S, R2 and R6 source sheets.",
    )
with right:
    st.markdown("### **02. Report Excel Sample**")
    report_file = st.file_uploader(
        "📂 Choose Report Excel file",
        type=["xlsx"],
        key="report_file",
        help="The original filled Report workbook is used as the output template.",
    )

st.markdown("<div class='section-label'>02 / GENERATE REPORT</div>", unsafe_allow_html=True)
st.markdown("**The original Report Excel will not be overwritten. A separate downloadable workbook will be created.**")
if pivot_file and report_file:
    st.success("Both files are ready. Select Generate Report to calculate the current rules.")
else:
    st.info("စတင်ရန် Auto_Pivot Excel နဲ့ Report Excel နမူနာဖိုင် နှစ်ဖိုင်လုံး Upload လုပ်ပါ။")

if st.button("🚀  GENERATE REPORT", type="primary", use_container_width=True, disabled=not (pivot_file and report_file)):
    try:
        pivot_bytes = pivot_file.getvalue()
        xls = pd.ExcelFile(io.BytesIO(pivot_bytes), engine="openpyxl")
        source_names = ["R1", "R1S", "R2", "R6"]
        missing = [name for name in source_names if name not in xls.sheet_names]
        if missing:
            st.error("Auto_Pivot ဖိုင်ထဲမှာ လိုအပ်တဲ့ sheet မတွေ့ပါ: " + ", ".join(missing))
            st.stop()

        frames = {name: pd.read_excel(xls, sheet_name=name) for name in source_names}
        # Map to the report labels requested by the user.
        labels = {"R1": "R1 (wo R1S)", "R1S": "R1S", "R2": "R2 (wo R2S)", "R6": "R6"}
        records = [summarize(frames[name], labels[name]) for name in source_names]
        metrics = pd.DataFrame(records)
        report_date = date_label(frames)
        result_bytes = build_report(report_file.getvalue(), metrics, report_date)

        st.markdown("<div class='section-label'>03 / CALCULATION PREVIEW</div>", unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Visited TKT", int(metrics["Visited TKT"].sum()))
        c2.metric("Resolved", int(metrics["Resolved"].sum()))
        c3.metric("Total Team", int(metrics["Total Team"].sum()))
        c4.metric("Remote Resolved (Auto)", int(metrics["Remote Resolved (Auto)"].sum()))
        st.dataframe(metrics, use_container_width=True, hide_index=True)
        st.warning(
            "မှတ်ချက် — လက်ရှိအဆင့်မှာ သတ်မှတ်ထားတဲ့ calculation summary ကို Report workbook ထဲက "
            "Automation Summary sheet အသစ်မှာ ထည့်ပေးပါတယ်။ Report Oct ရဲ့ မူရင်း sheet တွေထဲက "
            "တိကျတဲ့ cell နေရာတွေကို အတည်မပြုရသေးတဲ့အတွက် မူရင်း cell တွေကို မပြောင်းထားပါ။ "
            "Remote Resolved (Auto) ကိုလည်း summary ထဲမှာ သီးခြားပြထားပါတယ်။"
        )
        filename = f"Daily_Report_Automation_{report_date}.xlsx"
        st.download_button(
            "⬇️  DOWNLOAD GENERATED REPORT",
            data=result_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.success("Report workbook ပြင်ဆင်ပြီးပါပြီ။ Download ကိုနှိပ်ပြီး သိမ်းနိုင်ပါပြီ။")
    except Exception as exc:
        st.error(f"Report ထုတ်ရာတွင် ပြဿနာရှိနေပါတယ်: {exc}")
        st.exception(exc)

st.markdown("---")
st.markdown(
    "<div class='footnote'>Daily Report Automation · Secure, consistent and management-ready reporting</div>",
    unsafe_allow_html=True,
)
