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
    labor_c = col(df, 'Labor')
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

    # One engineer shared by R2 and R2S is assigned to only one side:
    # whichever side has more KPI-Resolved tickets. Ties stay with R2.
    staff = d[~d['_repay'].str.contains(r'engr\s*leave|route\s*cancel', regex=True, na=False)]
    owner = {}
    for eng, rows in staff[staff['_eng'].ne('')].groupby('_eng'):
        sides = set(rows['_area'])
        if len(sides) == 1:
            owner[eng] = next(iter(sides))
            continue
        counts = rows.assign(_is_resolved=rows['_status'].isin(RESOLVED)).groupby('_area')['_is_resolved'].sum().to_dict()
        owner[eng] = 'R2S' if counts.get('R2S', 0) > counts.get('R2', 0) else 'R2'
    d['_eng_owner'] = d['_eng'].map(owner).fillna('')

    # Assign each Car ID found in both R2 and R2S to only one side.
    # Compare KPI-Resolved rows for that Car ID by side; ties stay with R2.
    car_owner = {}
    car_rows = staff[staff['_car'].ne('')]
    for car, rows in car_rows.groupby('_car'):
        sides = set(rows['_area'])
        if len(sides) == 1:
            car_owner[car] = next(iter(sides))
            continue
        counts = rows.assign(_is_resolved=rows['_status'].isin(RESOLVED)).groupby('_area')['_is_resolved'].sum().to_dict()
        car_owner[car] = 'R2S' if counts.get('R2S', 0) > counts.get('R2', 0) else 'R2'
    d['_car_owner'] = d['_car'].map(car_owner).fillna('')
    return d


def area_metrics(df):
    status = df['_status']
    tkt = df['_kind'].eq('TKT')
    visited = int((tkt & status.isin(VISITED_TKT)).sum())
    resolved = int(status.isin(RESOLVED).sum())
    remote = int(status.eq('resolved (auto)').sum())
    unresolved = int(status.str.contains('unresolved', na=False).sum())
    # Unique engineer count ignores blank names; only rows with Task No were retained.
    staff = df[~df['_repay'].str.contains(r'engr\s*leave|route\s*cancel', regex=True, na=False)]
    if '_area' in df.columns and '_eng_owner' in df.columns:
        side = df['_area'].iloc[0] if len(df) and df['_area'].nunique() == 1 else None
        if side:
            staff_engs = staff[staff['_eng_owner'].eq(side)]
        else:
            staff_engs = staff
        engs = set(staff_engs['_eng'].replace('', pd.NA).dropna())
    else:
        engs = set(staff['_eng'].replace('', pd.NA).dropna())
    if '_area' in df.columns and '_car_owner' in df.columns:
        side = df['_area'].iloc[0] if len(df) and df['_area'].nunique() == 1 else None
        if side:
            staff_cars = staff[staff['_car_owner'].eq(side)]
        else:
            staff_cars = staff
        cars = set(staff_cars['_car'].replace('', pd.NA).dropna())
    else:
        cars = set(staff['_car'].replace('', pd.NA).dropna())
    by_class = {}
    for cat in ['FR-SLA / Biz TKT', 'Other TKT', 'FT-SBS', 'Install Express', 'Install POI']:
        subset = df[df['_class'].eq(cat)]
        by_class[cat] = {
            'blog': int(len(subset)),
            'visited': int((subset['_kind'].eq('TKT') & subset['_status'].isin(VISITED_TKT)).sum()),
            'unresolved': int(subset['_status'].str.contains('unresolved', na=False).sum()),
            'resolved': int(subset['_status'].isin(RESOLVED).sum()),
            'remote': int(subset['_status'].eq('resolved (auto)').sum()),
        }
    return {'blog': len(df), 'visited': visited, 'unresolved': unresolved,
            'resolved': resolved, 'remote': remote, 'eng': len(engs), 'car': len(cars), 'labor': len(labor),
            'by_class': by_class}


def fill_area(ws, df, start_col, rows, categories, team_row, npt=False):
    m = area_metrics(df)
    # Categories are mapped to the visible report's category columns in order.
    for offset, cat in enumerate(categories):
        c = start_col + offset
        values = [m['by_class'][cat][k] for k in ('blog', 'visited', 'unresolved', 'resolved', 'remote')]
        for rr, value in zip(rows, values):
            ws.cell(rr, c).value = value
    # Team/Car/Labor/Resolved row. Each distinct nonblank Labor entry counts once;
    # e.g. C-12/C-13 is one Labor entry, not two.
    ws.cell(team_row, start_col).value = m['eng']
    ws.cell(team_row, start_col + 1).value = m['car']
    ws.cell(team_row, start_col + 2).value = m['labor']
    ws.cell(team_row, start_col + 3).value = m['resolved']




def car_type_metrics(df, area=None):
    """Return Cars, distinct Teams, and KPI-resolved TKT+POI by car-sharing type."""
    d = df.copy()
    if area is not None and '_area' in d.columns:
        d = d[d['_area'].eq(area)].copy()
        if '_car_owner' in d.columns:
            d = d[d['_car_owner'].eq(area)].copy()

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

    # Naypyitaw is specifically sourced from R6 OC where Township = Zabuthiri.
    r6oc_name = next((s for s in ['R6 OC', 'R6OC'] if s in xls.sheet_names), None)
    if r6oc_name:
        r6oc_raw = pd.read_excel(xls, sheet_name=r6oc_name, dtype=object)
        r6oc = prep_frame(r6oc_raw, 'R6 OC')
        r6oc = r6oc[r6oc['_town'].map(lambda t: t == 'zabuthiri' or t.startswith('zabuthiri '))]
    else:
        r6oc = pd.DataFrame(columns=['_task','_status','_eng','_car','_labor','_town','_service','_root','_repay','_kind','_class','_task_present'])

    wb = load_workbook(io.BytesIO(report_bytes))
    if 'New CEO Lan Report' not in wb.sheetnames:
        raise ValueError("Report Excel ထဲမှာ 'New CEO Lan Report' Sheet မတွေ့ပါ")
    ws = wb['New CEO Lan Report']
    # Report's existing layout; only write data cells in the known report tables.
    # R1 table B:F, rows 39:43; R2 table N:R, rows 39:43.
    # R1S table B:F, rows 53:57; R2S table N:R, rows 53:57.
    # MDY table B:E, rows 68:72; Naypyitaw table N:Q, rows 68:72.
    fill_area(ws, frames['R1'], 2, [39,40,41,42,43], ['FR-SLA / Biz TKT','Other TKT','FT-SBS','Install Express','Install POI'], 44)
    fill_area(ws, r2[r2['_area'].eq('R2')], 14, [39,40,41,42,43], ['FR-SLA / Biz TKT','Other TKT','FT-SBS','Install Express','Install POI'], 44)
    fill_area(ws, frames['R1S'], 2, [53,54,55,56,57], ['FR-SLA / Biz TKT','Other TKT','FT-SBS','Install Express','Install POI'], 58)
    fill_area(ws, r2[r2['_area'].eq('R2S')], 14, [53,54,55,56,57], ['FR-SLA / Biz TKT','Other TKT','FT-SBS','Install Express','Install POI'], 58)
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
        'R6Oth': r6oc,
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
        'R2': area_metrics(r2[r2['_area'].eq('R2')]), 'R2S': area_metrics(r2[r2['_area'].eq('R2S')]),
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
.block-container{max-width:1240px;padding-top:1.2rem;padding-bottom:2.5rem}
.stApp{background:#F1F4F8;color:#202938}
.hero{background:linear-gradient(120deg,#101827,#26364D);color:#fff;padding:30px 34px;border-radius:18px;margin-bottom:22px}
.hero h1,.hero p{color:#fff!important;font-weight:900}
.hero h1{font-size:clamp(26px,3vw,36px);margin:10px 0}
.hero p{font-size:16px}
.section{font-weight:900;letter-spacing:1px;color:#26364D;margin:24px 0 12px}
[data-testid="stFileUploader"]{background:#fff;border:1px solid #C7D2E0;border-radius:14px;padding:14px}
[data-testid="stFileUploader"] label,[data-testid="stFileUploaderDropzone"] *{color:#172033!important;font-weight:800!important}
div.stButton>button,div.stDownloadButton>button{min-height:50px;font-weight:900;border-radius:11px}
</style>
<div class="hero"><div style="font-size:12px;letter-spacing:2px;font-weight:900">REPORTING WORKSPACE</div><h1>📊 DAILY REPORT AUTOMATION</h1><p>Automated Excel Reporting System · Upload → Calculate → Download</p></div>
''', unsafe_allow_html=True)

st.markdown('<div class="section">01 / UPLOAD FILES</div>', unsafe_allow_html=True)
a, b = st.columns(2, gap='large')
with a:
    st.markdown('### Auto_Pivot Excel')
    pivot_file = st.file_uploader('Auto_Pivot workbook ကိုရွေးပါ', type=['xlsx'], key='pivot')
with b:
    st.markdown('### Report Excel Sample')
    report_file = st.file_uploader('ဖြည့်စေချင်တဲ့ Report Excel ကိုရွေးပါ', type=['xlsx'], key='report')
st.caption('မူရင်း Report ကို overwrite မလုပ်ပါ။ ဖြည့်ပြီးသား workbook အသစ်ကို Download လုပ်ပေးပါမယ်။')

st.markdown('<div class="section">02 / GENERATE REPORT</div>', unsafe_allow_html=True)
if st.button('🚀 GENERATE REPORT', type='primary', use_container_width=True, disabled=not (pivot_file and report_file)):
    try:
        with st.spinner('Auto_Pivot Data ကိုတွက်ပြီး မင်းတင်ထားတဲ့ Report Excel ထဲ ဖြည့်နေပါတယ်…'):
            result, metrics = calculate_and_fill(pivot_file.getvalue(), report_file.getvalue())
        st.success('✅ Report Excel ထဲကို Data ဖြည့်ပြီးပါပြီ။')
        st.markdown('### Calculation Preview')
        preview = []
        for area, m in metrics.items():
            preview.append({'Area': area, 'Total Blog': m['blog'], 'Visited TKT': m['visited'], 'Unresolved': m['unresolved'], 'Resolved': m['resolved'], 'Remote Resolved': m['remote'], 'Eng': m['eng'], 'Car ID': m['car'], 'Labor': m['labor']})
        st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)
        st.download_button('⬇️ DOWNLOAD FILLED REPORT EXCEL', data=result, file_name='Daily_Report_Filled.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary', use_container_width=True)
    except Exception as e:
        st.error(f'Report ဖြည့်ရာတွင် Error ဖြစ်နေပါတယ်: {e}')
