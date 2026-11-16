import io, os, re
from datetime import date
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title='Daily Report Automation', page_icon='📊', layout='wide')

# Set APP_PASSWORD in Streamlit Secrets; never hard-code the password here.
APP_PASSWORD = st.secrets.get('APP_PASSWORD', os.getenv('APP_PASSWORD', ''))
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

# ---------------- Report rules ----------------
RESOLVED = {'resolved', 'resolved (no kpi)', 'resolved (ap up)', 'resolved (onu up)'}
VISITED_TKT = {
    'engr visited pending', 'handover to fttx', 'handover to qc-wireless',
    'handover to s-combine', 'handover to sales', 'resolved',
    'resolved (no kpi)', 'resolved (ap up)', 'resolved (onu up)',
    'unresolved(cabling done)'
}
R2S_TOWNSHIPS = {
    'kamaryut', 'dagon myothit north', 'sanchaung', 'kyeemyindaing',
    'lanmadaw', 'pabedan', 'pazundaung', 'kyauktada', 'latha', 'botahtaung'
}


def clean(v):
    if v is None or pd.isna(v):
        return ''
    return str(v).strip()


def key(v):
    return re.sub(r'\s+', ' ', clean(v)).casefold()


def col(df, name):
    target = key(name)
    return next((c for c in df.columns if key(c) == target), None)


def task_kind(v):
    t = clean(v).upper()
    if 'TKT' in t:
        return 'TKT'
    if 'POI' in t:
        return 'POI'
    return ''


def classification(row):
    task = task_kind(row.get('_task', ''))
    service = clean(row.get('_service', '')).upper()
    root = clean(row.get('_root', '')).upper()
    if task == 'TKT':
        if 'FR-SLA' in service or 'BIZ' in service or 'INSTALLATION TYPE CHANGE' in root:
            return 'FR-SLA / Biz TKT'
        if 'FT-SBS' in root:
            return 'FT-SBS'
        return 'Other TKT'
    if task == 'POI':
        return 'Install Express' if 'EXPRESS' in root else 'Install POI'
    return ''


def prep_frame(df, area):
    df = df.dropna(how='all').copy()
    task_c = col(df, 'Task No')
    status_c = col(df, 'Task Status')
    eng_c = col(df, 'LAN Engineer Name')
    car_c = col(df, 'Car ID')
    labor_c = col(df, 'Contractor / Office Labour ID')
    town_c = col(df, 'Township')
    service_c = col(df, 'Service Area')
    root_c = col(df, 'Root Cause')
    repay_c = col(df, 'Repay Status')
    out = pd.DataFrame(index=df.index)
    out['_task'] = df[task_c].map(clean) if task_c else ''
    out['_status'] = df[status_c].map(key) if status_c else ''
    out['_eng'] = df[eng_c].map(clean) if eng_c else ''
    out['_car'] = df[car_c].map(clean) if car_c else ''
    out['_labor'] = df[labor_c].map(clean) if labor_c else ''
    out['_town'] = df[town_c].map(key) if town_c else ''
    out['_service'] = df[service_c].map(clean) if service_c else ''
    out['_root'] = df[root_c].map(clean) if root_c else ''
    out['_repay'] = df[repay_c].map(key) if repay_c else ''
    out['_kind'] = out['_task'].map(task_kind)
    out['_class'] = out.apply(classification, axis=1)
    out['_task_present'] = out['_task'].ne('')
    out = out[out['_task_present']].copy()
    # Exclude rows explicitly marked as route cancelled from staffing/car counts.
    return out


def split_r2(df):
    d = df.copy()

    def is_r2s(t):
        base = re.split(r'\s*\(|\s+ward\b', t, maxsplit=1)[0].strip()
        return t in R2S_TOWNSHIPS or base in R2S_TOWNSHIPS

    d['_area'] = d['_town'].map(lambda t: 'R2S' if is_r2s(t) else 'R2')
    # Staffing and car ownership ignore leave/cancelled rows.
    staff = d[~d['_repay'].str.contains(
        r'engr\s*leave|route\s*cancel', regex=True, na=False
    )]

    # Each Car ID is one indivisible group, including One Car.
    # Assign the complete group to the side with more KPI-Resolved rows.
    # A tie is assigned to R2S.
    car_owner = {}
    for car, rows in staff[staff['_car'].ne('')].groupby('_car'):
        counts = rows.assign(
            _is_resolved=rows['_status'].isin(RESOLVED)
        ).groupby('_area')['_is_resolved'].sum().to_dict()
        if len(set(rows['_area'])) == 1:
            car_owner[car] = next(iter(set(rows['_area'])))
        else:
            car_owner[car] = (
                'R2S' if counts.get('R2S', 0) >= counts.get('R2', 0) else 'R2'
            )
    d['_car_owner'] = d['_car'].map(car_owner).fillna('')

    # Engineers with a Car ID follow their Car ID's owner. Engineers without
    # a Car ID are assigned individually by KPI-Resolved count; ties go R2S.
    eng_owner = {}
    with_car = staff[staff['_eng'].ne('') & staff['_car'].ne('')]
    for eng, rows in with_car.groupby('_eng'):
        owners = rows['_car'].map(car_owner).dropna()
        owners = owners[owners.ne('')]
        if not owners.empty:
            counts = owners.value_counts()
            eng_owner[eng] = (
                'R2S' if counts.get('R2S', 0) >= counts.get('R2', 0) else 'R2'
            )

    no_car = staff[staff['_eng'].ne('') & staff['_car'].eq('')]
    for eng, rows in no_car.groupby('_eng'):
        counts = rows.assign(
            _is_resolved=rows['_status'].isin(RESOLVED)
        ).groupby('_area')['_is_resolved'].sum().to_dict()
        if len(set(rows['_area'])) == 1:
            eng_owner[eng] = next(iter(set(rows['_area'])))
        else:
            eng_owner[eng] = (
                'R2S' if counts.get('R2S', 0) >= counts.get('R2', 0) else 'R2'
            )
    d['_eng_owner'] = d['_eng'].map(eng_owner).fillna('')

    # One common owner field drives R2/R2S Resolved, HC, Labor and Car KPIs.
    # Car ID rows follow the Car ID owner; no-Car rows follow the Engineer owner.
    d['_owner_area'] = d['_area']
    has_car_owner = d['_car'].ne('') & d['_car_owner'].ne('')
    d.loc[has_car_owner, '_owner_area'] = d.loc[has_car_owner, '_car_owner']
    no_car_owner = d['_car'].eq('') & d['_eng_owner'].ne('')
    d.loc[no_car_owner, '_owner_area'] = d.loc[no_car_owner, '_eng_owner']
    return d


def area_metrics(df, owner_area=None):
    # Blog / Visited / Unresolved remain with original Township.
    source_df = df
    if owner_area is not None and '_area' in df.columns:
        source_df = df[df['_area'].eq(owner_area)]

    status = source_df['_status']
    visited = int((source_df['_kind'].isin(['TKT', 'POI']) & status.isin(VISITED_TKT)).sum())
    remote = int(status.eq('resolved (auto)').sum())
    unresolved = int(status.str.contains('unresolved', na=False).sum())

    # Resolved and staffing metrics follow the shared Car ID / Engineer owner.
    owner_df = df
    if owner_area is not None and '_owner_area' in df.columns:
        owner_df = df[df['_owner_area'].eq(owner_area)]
    resolved = int(owner_df['_status'].isin(RESOLVED).sum())

    staff = owner_df[~owner_df['_repay'].str.contains(
        r'engr\s*leave|route\s*cancel', regex=True, na=False
    )]
    engs = set(staff['_eng'].replace('', pd.NA).dropna())
    cars = set(staff['_car'].replace('', pd.NA).dropna())
    labor_values = staff['_labor'].map(clean) if '_labor' in staff.columns else pd.Series(dtype=object)
    labor = {
        key(value) for value in labor_values
        if value and key(value) not in {'no cabler', 'no labor', 'no labour'}
    }

    by_class = {}
    for cat in ['FR-SLA / Biz TKT', 'Other TKT', 'FT-SBS', 'Install Express', 'Install POI']:
        subset = source_df[source_df['_class'].eq(cat)]
        resolved_subset = owner_df[owner_df['_class'].eq(cat)]
        by_class[cat] = {
            'blog': int(len(subset)),
            'visited': int((subset['_kind'].isin(['TKT', 'POI']) & subset['_status'].isin(VISITED_TKT)).sum()),
            'unresolved': int(subset['_status'].str.contains('unresolved', na=False).sum()),
            'resolved': int(resolved_subset['_status'].isin(RESOLVED).sum()),
            'remote': int(subset['_status'].eq('resolved (auto)').sum()),
        }
    return {'blog': len(source_df), 'visited': visited, 'unresolved': unresolved,
            'resolved': resolved, 'remote': remote, 'eng': len(engs),
            'car': len(cars), 'labor': len(labor), 'by_class': by_class}


def fill_area(ws, df, start_col, rows, categories, team_row, npt=False, owner_area=None):
    m = area_metrics(df, owner_area=owner_area)
    # Categories are mapped to the visible report's category columns in order.
    for offset, cat in enumerate(categories):
        c = start_col + offset
        values = [m['by_class'][cat][k] for k in ('blog', 'visited', 'unresolved', 'resolved', 'remote')]
        for row_index, (rr, value) in enumerate(zip(rows, values)):
            # Preserve Blog values from the uploaded Report Excel; update other metrics only.
            if row_index == 0:
                continue
            ws.cell(rr, c).value = value
    # Naypyitaw (npt=True) must not collect or write Labor; preserve its Eng/Car/Resolved layout.
    ws.cell(team_row, start_col).value = m['eng']
    ws.cell(team_row, start_col + 1).value = m['car']
    if npt:
        ws.cell(team_row, start_col + 2).value = m['resolved']
    else:
        # Each distinct nonblank Contractor / Office Labour ID entry counts once.
        ws.cell(team_row, start_col + 2).value = m['labor']
        ws.cell(team_row, start_col + 3).value = m['resolved']




def car_type_metrics(df, area=None):
    """Return Cars, distinct Teams, and KPI-resolved TKT+POI by car-sharing type."""
    d = df.copy()
    if area is not None and '_area' in d.columns:
        if '_owner_area' in d.columns:
            d = d[d['_owner_area'].eq(area)].copy()
        else:
            d = d[d['_area'].eq(area)].copy()

    # Exclude staffing/car rows marked as Engr Leave or Route Cancel.
    staff = d[
        ~d['_repay'].str.contains(r'engr\s*leave|route\s*cancel', regex=True, na=False)
        & d['_car'].ne('')
        & d['_eng'].ne('')
    ].copy()

    if staff.empty:
        return {1: (0, 0, 0), 2: (0, 0, 0), 3: (0, 0, 0)}

    # Car type is based on the number of distinct engineers assigned to each Car ID.
    engineers_by_car = staff.groupby('_car')['_eng'].agg(lambda s: set(s.dropna()) - {''}).to_dict()
    car_type = {car: len(engs) for car, engs in engineers_by_car.items()}

    # Resolved TKT + POI: only KPI-resolved statuses and Task No containing TKT or POI.
    resolved_rows = d[
        d['_status'].isin(RESOLVED)
        & d['_kind'].isin(['TKT', 'POI'])
        & d['_car'].ne('')
    ].copy()
    resolved_by_car = resolved_rows.groupby('_car').size().to_dict()

    results = {}
    for n in (1, 2, 3):
        selected = {car for car, count in car_type.items() if count == n}
        teams = set()
        for car in selected:
            teams.update(engineers_by_car.get(car, set()))
        resolved_count = sum(int(resolved_by_car.get(car, 0)) for car in selected)
        results[n] = (len(selected), len(teams), resolved_count)
    return results


def fill_car_type_kpi(ws, areas):
    """Fill only the numeric cells in the existing 3 Car Type KPI table (rows 21:31)."""
    # Table columns B:H: R1, R1S, R2, R2S, MDY, R1Oth, R6Oth.
    # Each area entry is (dataframe, split-area-or-None).
    col_map = [
        ('R1', 2, None),
        ('R1S', 3, None),
        ('R2', 4, 'R2'),
        ('R2S', 5, 'R2S'),
        ('MDY', 6, None),
        ('R1Oth', 7, None),
        ('R6Oth', 8, None),
    ]
    for name, col_num, split_area in col_map:
        df = areas.get(name)
        if df is None:
            continue
        metrics = car_type_metrics(df, split_area)
        for n, rows in [(1, (21, 22, 23)), (2, (25, 26, 27)), (3, (29, 30, 31))]:
            cars, teams, resolved = metrics[n]
            for row_num, value in zip(rows, (cars, teams, resolved)):
                ws.cell(row_num, col_num).value = value



def fill_outer_city_kpi(ws, xls, r1oc_name=None, r6oc_name=None):
    """Fill only TKT Resolved, POI Resolved, and Teams for the seven outer-city rows."""
    city_rows = [
        ('Bago', 10, r1oc_name),
        ('Meikhtila', 11, r1oc_name),
        ('Taungoo', 12, r1oc_name),
        ('Magway', 13, r1oc_name),
        ('Taunggyi', 14, r6oc_name),
        ('Mawlamyine', 15, r6oc_name),
        ('Pyinmana', 16, r6oc_name),
    ]
    cache = {}
    for city, row_num, sheet_name in city_rows:
        if not sheet_name:
            continue
        if sheet_name not in cache:
            cache[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name, dtype=object)
        raw = cache[sheet_name]
        city_col = col(raw, 'City')
        task_col = col(raw, 'Task No')
        status_col = col(raw, 'Task Status')
        eng_col = col(raw, 'LAN Engineer Name')
        repay_col = col(raw, 'Repay Status')
        if city_col is None:
            continue

        subset = raw[raw[city_col].map(lambda v: key(v) == key(city))].copy()
        tkt_resolved = 0
        poi_resolved = 0
        engineers = set()
        for _, record in subset.iterrows():
            task = clean(record[task_col]) if task_col else ''
            status = key(record[status_col]) if status_col else ''
            eng = clean(record[eng_col]) if eng_col else ''
            repay = key(record[repay_col]) if repay_col else ''
            if task and status in RESOLVED:
                kind = task_kind(task)
                if kind == 'TKT':
                    tkt_resolved += 1
                elif kind == 'POI':
                    poi_resolved += 1
            if eng and repay not in {'engr leave', 'route cancel', 'return cancelled', 'permanent cancelled'}:
                engineers.add(eng)

        # Existing report columns: D = TKT Resolved, F = POI Resolved, H = Teams.
        ws.cell(row_num, 4).value = tkt_resolved
        ws.cell(row_num, 6).value = poi_resolved
        ws.cell(row_num, 8).value = len(engineers)


def calculate_and_fill(pivot_bytes, report_bytes):
    xls = pd.ExcelFile(io.BytesIO(pivot_bytes), engine='openpyxl')
    # Raw source sheets are taken from the uploaded Auto_Pivot workbook, not its PivotTable summary sheets.
    required = ['R1', 'R1S', 'R2', 'R6']
    missing = [s for s in required if s not in xls.sheet_names]
    if missing:
        raise ValueError('Auto_Pivot ထဲမှာ Source Sheet မတွေ့ပါ: ' + ', '.join(missing))
    raw = {s: pd.read_excel(xls, sheet_name=s, dtype=object) for s in required}
    frames = {s: prep_frame(df, s) for s, df in raw.items()}
    r2 = split_r2(frames['R2'])

    # Keep the complete R6 OC source for the R6Oth 3 Car Type KPI.
    # Naypyitaw's report section uses only the Zabuthiri subset.
    r6oc_name = next((s for s in ['R6 OC', 'R6OC'] if s in xls.sheet_names), None)
    if r6oc_name:
        r6oc_raw = pd.read_excel(xls, sheet_name=r6oc_name, dtype=object)
        r6oc_all = prep_frame(r6oc_raw, 'R6 OC')
        r6oc = r6oc_all[r6oc_all['_town'].map(lambda t: t == 'zabuthiri' or t.startswith('zabuthiri '))].copy()
    else:
        empty_columns = ['_task','_status','_eng','_car','_labor','_town','_service','_root','_repay','_kind','_class','_task_present']
        r6oc_all = pd.DataFrame(columns=empty_columns)
        r6oc = pd.DataFrame(columns=empty_columns)

    wb = load_workbook(io.BytesIO(report_bytes))
    if 'New CEO Lan Report' not in wb.sheetnames:
        raise ValueError("Report Excel ထဲမှာ 'New CEO Lan Report' Sheet မတွေ့ပါ")
    ws = wb['New CEO Lan Report']
    # Report's existing layout; only write data cells in the known report tables.
    # R1 table B:F, rows 39:43; R2 table N:R, rows 39:43.
    # R1S table B:F, rows 53:57; R2S table N:R, rows 53:57.
    # MDY table B:E, rows 68:72; Naypyitaw table N:Q, rows 68:72.
    fill_area(ws, frames['R1'], 2, [39,40,41,42,43], ['FR-SLA / Biz TKT','Other TKT','FT-SBS','Install Express','Install POI'], 44)
    fill_area(ws, r2, 14, [39,40,41,42,43], ['FR-SLA / Biz TKT','Other TKT','FT-SBS','Install Express','Install POI'], 44, owner_area='R2')
    fill_area(ws, frames['R1S'], 2, [53,54,55,56,57], ['FR-SLA / Biz TKT','Other TKT','FT-SBS','Install Express','Install POI'], 58)
    fill_area(ws, r2, 14, [53,54,55,56,57], ['FR-SLA / Biz TKT','Other TKT','FT-SBS','Install Express','Install POI'], 58, owner_area='R2S')
    fill_area(ws, frames['R6'], 2, [68,69,70,71,72], ['FR-SLA / Biz TKT','Other TKT','Install Express','Install POI'], 73)
    fill_area(ws, r6oc, 14, [68,69,70,71,72], ['FR-SLA / Biz TKT','Other TKT','Install Express','Install POI'], 73, npt=True)
    # Add-on: fill the existing 3 Car Type KPI table only; other report cells remain unchanged.
    r1oc_name = next((s for s in ['R1 OC', 'R1OC'] if s in xls.sheet_names), None)
    r6oc_name = next((s for s in ['R6 OC', 'R6OC'] if s in xls.sheet_names), None)
    if r1oc_name:
        r1oc_raw = pd.read_excel(xls, sheet_name=r1oc_name, dtype=object)
        r1oc = prep_frame(r1oc_raw, 'R1 OC')
    else:
        r1oc = None

    car_type_areas = {
        'R1': frames['R1'],
        'R1S': frames['R1S'],
        'R2': r2,
        'R2S': r2,
        'MDY': frames['R6'],
        'R1Oth': r1oc,
        'R6Oth': r6oc_all,
    }
    fill_car_type_kpi(ws, car_type_areas)
    # Add-on: fill only outer-city TKT Resolved, POI Resolved, and Teams cells.
    fill_outer_city_kpi(ws, xls, r1oc_name, r6oc_name)

    # Set date cells using today's date; formula cells in the report remain intact.
    report_date = date.today().strftime('%d-%b')
    for cell in ['A37','M37','A51','M51','A66','M66']:
        ws[cell] = report_date
    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue(), {
        'R1': area_metrics(frames['R1']), 'R1S': area_metrics(frames['R1S']),
        'R2': area_metrics(r2, owner_area='R2'), 'R2S': area_metrics(r2, owner_area='R2S'),
        'MDY': area_metrics(frames['R6']), 'Naypyitaw': area_metrics(r6oc)
    }

# ---------------- UI ----------------
if APP_PASSWORD:
    if not st.session_state.authenticated:
        st.title('🔐 Daily Report Automation')
        st.caption('Secure access · Daily Reporting System')
        pw = st.text_input('Password', type='password')
        if st.button('LOGIN', type='primary', use_container_width=True):
            if pw == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error('Password မမှန်ပါ။')
        st.stop()

st.markdown('''
<style>
.block-container{max-width:1440px;padding-top:1.15rem;padding-bottom:3rem}
.stApp{background:#F1F4F8;color:#202938}
.hero{background:linear-gradient(120deg,#101827 0%,#26364D 100%);color:#fff;padding:25px 32px;border-radius:18px;margin-bottom:24px;box-shadow:0 8px 24px rgba(15,23,42,.10)}
.hero h1,.hero p{color:#fff!important;font-weight:850}
.hero h1{font-size:clamp(25px,3vw,35px);margin:7px 0 9px}
.hero p{font-size:15px;margin:0;color:#E2E8F0!important}
.hero-kicker{font-size:11px;letter-spacing:2px;font-weight:850;color:#B8C7DA}
.section{font-weight:900;letter-spacing:1px;color:#26364D;margin:24px 0 12px}
[data-testid="stFileUploader"]{background:#fff;border:1px solid #CBD5E1;border-radius:14px;padding:14px;box-shadow:0 2px 8px rgba(15,23,42,.035)}
[data-testid="stFileUploader"] label{color:#172033!important;font-weight:750!important}
[data-testid="stFileUploaderDropzone"]{background:#F8FAFC!important;border:1px dashed #94A3B8!important;border-radius:10px!important}
[data-testid="stFileUploaderDropzone"] *{color:#334155!important}
div.stButton>button,div.stDownloadButton>button{min-height:48px;font-weight:800;border-radius:11px;border:0}
div.stButton>button[kind="primary"]{background:linear-gradient(90deg,#1D4ED8,#2563EB);color:#fff}
div.stButton>button[kind="primary"]:hover{background:#1E40AF;color:#fff}
[data-testid="stTabs"] button{font-weight:750;color:#334155}
[data-testid="stDataFrame"]{border:1px solid #D5DDE7;border-radius:10px;overflow:hidden}
[data-testid="stStatusWidget"]{background:#fff;border:1px solid #CBD5E1;border-radius:12px;padding:12px}
.section{font-size:1.12rem;font-weight:900;letter-spacing:.6px;color:#142B4A!important;margin:26px 0 13px;padding:10px 14px;background:#FFFFFF;border-left:5px solid #2563EB;border-radius:8px;box-shadow:0 2px 7px rgba(15,23,42,.045)}
.preview-title{color:#142B4A!important;font-weight:850!important;margin:8px 0 12px}
[data-testid="stTabs"] button{font-weight:800!important;color:#FFFFFF!important;background:#142B4A!important;border-radius:8px 8px 0 0;margin-right:4px}
[data-testid="stTabs"] button[aria-selected="true"]{background:#2563EB!important;color:#FFFFFF!important}
.excel-preview-scroll{width:100%;max-height:680px;overflow:auto;background:#FFFFFF;border:1px solid #CBD5E1;border-radius:10px;box-shadow:0 3px 12px rgba(15,23,42,.06)}
table.excel-preview{border-collapse:collapse;table-layout:fixed;width:max-content;min-width:100%;font-family:Calibri,Arial,sans-serif;font-size:12px;color:#263445}
table.excel-preview td{border:1px solid #D5DDE7;min-width:42px;max-width:360px;overflow-wrap:anywhere}
table.excel-preview .rownum{position:sticky;left:0;z-index:2;background:#EEF2F7;color:#64748B;text-align:center;font-weight:600;border:1px solid #CBD5E1;min-width:42px}

.helper{color:#64748B;font-size:13px;margin-top:-4px}
</style>
<div class="hero">
  <div class="hero-kicker">REPORTING WORKSPACE</div>
  <h1>📊 DAILY REPORT AUTOMATION</h1>
  <p>Professional Reporting Dashboard&nbsp; · &nbsp;Upload → Process → Validate → Report</p>
</div>
''', unsafe_allow_html=True)

st.markdown('<div class="section">01 / UPLOAD FILES</div>', unsafe_allow_html=True)
a, b = st.columns(2, gap='large')
with a:
    st.markdown('### 📗 Auto_Pivot Excel')
    st.caption('Prepared Auto_Pivot workbook')
    pivot_file = st.file_uploader('Auto_Pivot workbook ကိုရွေးပါ', type=['xlsx'], key='pivot')
with b:
    st.markdown('### 📘 Report Excel Sample')
    st.caption('ဖြည့်စေချင်တဲ့ မူရင်း Report workbook')
    report_file = st.file_uploader('ဖြည့်စေချင်တဲ့ Report Excel ကိုရွေးပါ', type=['xlsx'], key='report')
st.caption('မူရင်း Report ကို overwrite မလုပ်ပါ။ ဖြည့်ပြီးသား workbook အသစ်ကို Download လုပ်ပေးပါမယ်။')

st.markdown('<div class="section">02 / GENERATE REPORT</div>', unsafe_allow_html=True)
if st.button('⚡ GENERATE DAILY REPORT', type='primary', use_container_width=True, disabled=not (pivot_file and report_file)):
    try:
        with st.status('Daily Report ကို ပြင်ဆင်နေပါသည်…', expanded=True) as process_status:
            st.write('① Auto_Pivot နှင့် Report Excel ဖိုင်များကို ဖတ်ရှုနေပါသည်။')
            st.write('② Region, Township နှင့် KPI များကို တွက်ချက်နေပါသည်။')
            result, metrics = calculate_and_fill(pivot_file.getvalue(), report_file.getvalue())
            st.write('③ တွက်ချက်ထားသော Data များကို Report Excel ထဲသို့ ဖြည့်သွင်းနေပါသည်။')
            st.write('④ Output workbook ကို ပြင်ဆင်ပြီးပါပြီ။')
            process_status.update(label='Report generation ပြီးဆုံးပါပြီ။', state='complete', expanded=False)
        st.session_state['daily_report_result'] = result
        st.session_state['daily_report_metrics'] = metrics
        st.session_state['daily_report_filename'] = 'Daily_Report_Filled.xlsx'
        st.success('✅ Daily Report ကို အောင်မြင်စွာ ပြင်ဆင်ပြီးပါပြီ။ အောက်တွင် Preview ကြည့်နိုင်ပါသည်။')
    except Exception as e:
        st.error(f'Report ဖြည့်ရာတွင် Error ဖြစ်နေပါတယ်: {e}')

# Excel-style preview renderer (UI only; does not modify the generated workbook).
def _excel_color(color, fallback):
    if color is None:
        return fallback
    try:
        if color.type == 'rgb' and color.rgb:
            value = str(color.rgb)[-6:]
            return '#' + value
        if color.type == 'indexed' and color.indexed is not None:
            indexed = {0: '#000000', 1: '#FFFFFF', 2: '#FF0000', 3: '#00FF00',
                       4: '#0000FF', 5: '#FFFF00', 6: '#FF00FF', 7: '#00FFFF',
                       8: '#000000', 9: '#FFFFFF'}
            return indexed.get(color.indexed, fallback)
    except Exception:
        pass
    return fallback


def _excel_preview_html(ws, cached_ws=None, min_row=1, max_row=None, min_col=1, max_col=None):
    import html
    from openpyxl.utils import get_column_letter

    merged_starts = {}
    merged_children = set()
    for merged in ws.merged_cells.ranges:
        m_min_col, m_min_row, m_max_col, m_max_row = merged.bounds
        merged_starts[(m_min_row, m_min_col)] = (m_max_row - m_min_row + 1, m_max_col - m_min_col + 1)
        for rr in range(m_min_row, m_max_row + 1):
            for cc in range(m_min_col, m_max_col + 1):
                if (rr, cc) != (m_min_row, m_min_col):
                    merged_children.add((rr, cc))

    last_row = min(max_row or ws.max_row, ws.max_row)
    last_col = min(max_col or ws.max_column, ws.max_column)
    first_row, first_col = max(1, min_row), max(1, min_col)
    while last_row > first_row and all(ws.cell(last_row, c).value is None for c in range(first_col, last_col + 1)):
        last_row -= 1
    while last_col > first_col and all(ws.cell(r, last_col).value is None for r in range(first_row, last_row + 1)):
        last_col -= 1

    colgroup = ['<col style="width:42px">']
    for c in range(first_col, last_col + 1):
        width = ws.column_dimensions[get_column_letter(c)].width
        px = max(42, min(360, int((width or 12) * 7)))
        colgroup.append(f'<col style="width:{px}px">')

    rows_html = []
    for r in range(first_row, last_row + 1):
        height = ws.row_dimensions[r].height
        row_style = f'height:{max(18, min(120, int(height * 1.35)))}px;' if height else ''
        cells = [f'<th class="rownum" style="{row_style}">{r}</th>']
        for c in range(first_col, last_col + 1):
            if (r, c) in merged_children:
                continue
            cell = ws.cell(r, c)
            value = cell.value
            if cached_ws is not None and isinstance(value, str) and value.startswith('='):
                cached_value = cached_ws.cell(r, c).value
                if cached_value is not None:
                    value = cached_value
            text_value = '' if value is None else html.escape(str(value)).replace('\\n', '<br>')
            style = []
            fill = cell.fill
            if fill and fill.fill_type:
                bg = _excel_color(fill.fgColor, '#FFFFFF')
                style.append(f'background:{bg}')
            font = cell.font
            style.append(f'color:{_excel_color(font.color, "#263445")}')
            style.append(f'font-size:{max(9, min(24, int(font.sz or 10)))}px')
            style.append('font-weight:700' if font.bold else 'font-weight:400')
            if font.italic:
                style.append('font-style:italic')
            if font.underline:
                style.append('text-decoration:underline')
            align = cell.alignment
            style.append('text-align:' + (align.horizontal or 'left'))
            style.append('vertical-align:' + ('middle' if align.vertical in (None, 'center') else align.vertical))
            style.append('white-space:pre-wrap')
            style.append('padding:4px 7px')
            for side_name in ('left', 'right', 'top', 'bottom'):
                side = getattr(cell.border, side_name)
                if side and side.style:
                    color = _excel_color(side.color, '#64748B')
                    border_width = '2px' if side.style in ('medium', 'thick', 'double') else '1px'
                    style.append(f'border-{side_name}:{border_width} solid {color}')
            if cell.number_format and isinstance(value, (int, float)):
                pass
            rowspan, colspan = merged_starts.get((r, c), (1, 1))
            spans = (f' rowspan="{rowspan}"' if rowspan > 1 else '') + (f' colspan="{colspan}"' if colspan > 1 else '')
            cells.append(f'<td{spans} style="{";".join(style)}">{text_value}</td>')
        rows_html.append('<tr>' + ''.join(cells) + '</tr>')

    return (
        '<div class="excel-preview-scroll"><table class="excel-preview"><colgroup>'
        + ''.join(colgroup) + '</colgroup><tbody>' + ''.join(rows_html)
        + '</tbody></table></div>'
    )


# Keep the generated workbook in session state so users can switch preview tabs
# without running the calculation again.
if st.session_state.get('daily_report_result'):
    result = st.session_state['daily_report_result']
    metrics = st.session_state.get('daily_report_metrics', {})
    st.markdown('<div class="section">03 / REPORT PREVIEW</div>', unsafe_allow_html=True)
    st.caption('မူရင်း Excel Sheet ရဲ့ Formatting နဲ့ Layout ကို ထိန်းသိမ်းပြီး Preview ပြထားပါသည်။')

    try:
        preview_wb = load_workbook(io.BytesIO(result), data_only=False, read_only=False)
        cached_wb = load_workbook(io.BytesIO(result), data_only=True, read_only=True)
        sheet_names = preview_wb.sheetnames
        main_sheet = 'New CEO Lan Report' if 'New CEO Lan Report' in sheet_names else sheet_names[0]

        def find_sheet(tokens):
            for sheet_name in sheet_names:
                normalized = sheet_name.casefold().replace('_', ' ').replace('-', ' ')
                if any(token in normalized for token in tokens):
                    return sheet_name
            return None

        regional_sheets = [name for name in sheet_names if name != main_sheet and any(
            token in name.casefold() for token in ['regional', 'region', 'detail kpi', 'outer city', 'r1 oc', 'r6 oc']
        )]

        tab_daily, tab_car, tab_region = st.tabs([
            '📋 Daily Report', '🚘 3 Car Type KPI', '🌏 Regional Reports'
        ])
        with tab_daily:
            st.markdown('<h3 class="preview-title">📋 Daily Report</h3>', unsafe_allow_html=True)
            st.caption(f'Source sheet: {main_sheet} · Excel-style worksheet preview')
            st.markdown(_excel_preview_html(preview_wb[main_sheet], cached_wb[main_sheet]), unsafe_allow_html=True)

        with tab_car:
            st.markdown('<h3 class="preview-title">🚘 3 Car Type KPI</h3>', unsafe_allow_html=True)
            st.caption('R1 · R1S · R2 · R2S · MDY · R1Oth · R6Oth')
            kpi_ws = preview_wb[main_sheet]
            st.markdown(_excel_preview_html(kpi_ws, cached_wb[main_sheet], min_row=20, max_row=32, min_col=1, max_col=8), unsafe_allow_html=True)

        with tab_region:
            st.markdown('<h3 class="preview-title">🌏 Regional Reports</h3>', unsafe_allow_html=True)
            if regional_sheets:
                for region_sheet in regional_sheets:
                    with st.expander(region_sheet, expanded=True):
                        st.markdown(_excel_preview_html(
                            preview_wb[region_sheet], cached_wb[region_sheet]
                        ), unsafe_allow_html=True)
            else:
                st.caption(f'Source sheet: {main_sheet} · Regional report sections')
                st.markdown(_excel_preview_html(
                    preview_wb[main_sheet], cached_wb[main_sheet], min_row=36, max_row=73
                ), unsafe_allow_html=True)
        preview_wb.close()
        cached_wb.close()
    except Exception as e:
        st.warning(f'Preview ပြသရာတွင် အခက်အခဲရှိနေပါသည်: {e}')

    st.markdown('<div class="section">04 / DOWNLOAD FINAL REPORT</div>', unsafe_allow_html=True)
    st.download_button(
        '⬇️ DOWNLOAD FILLED REPORT EXCEL', data=result,
        file_name='Daily_Report_Filled.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type='primary', use_container_width=True
    )
