# %%
import os, pathlib, json, re, requests, time
import pandas as pd
import polars as pl
from datetime import datetime, timedelta, time as dtime

TEAMS_WEBHOOK_URL = "https://default599e51d62f8c43478e591f795a51a9.8c.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/c24f30c010df45a6a6dac9421643bb34/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=5vWDl18a7-IWSvHuZAWgGtQcwM54nEapSArj4JVPnGg"

DATA_DIR      = r"C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\Rawdata\CAPTURE\current_agent"
SCHEDULE_FILE = r"C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\Schedule\Schedule (Ops version)\2026\Master_Schedule_Merged.xlsx"

LOB_MAP = {
    "LG Chat": [
        "Chat_OD_EN_Car_Activity", "Chat_OD_EN_Lodging",
        "Chat - Global English Lodging Nesting", "Chat_Lodging English w Car",
        "Chat_AC_GLB_EN_Lodging_Proficient", "Chat_AC_GLB_EN_Car_Activity",
        "Chat_AC_GLB_EN_Lodging_Expert",
    ],
    "NL Chat": [
        "Chat - Global English Non- Lodging Nesting", "Chat_OD_EN_Dual_GDS",
        "Chat_AC_GLB_EN_Proficient", "Chat_AC_GLB_EN_Expert",
    ],
}

LEAVE_CODES    = {'AL','LWP','CO','SL','EL','ML','PL','SPL','BL','CL'}
EXCLUDE_SHIFTS = {'OFF','TERMINATION','TERM','RESIGNED'}

LOC_STYLE = {
    "HCM":   {"bg": "#DA251D", "fg": "#FFD700"},
    "PUN":   {"bg": "#388E3C", "fg": "#ffffff"},
    "KOL":   {"bg": "#1565C0", "fg": "#ffffff"},
    "CAI":   {"bg": "#E65100", "fg": "#ffffff"},
    "OTHER": {"bg": "#757575", "fg": "#ffffff"},
}
LOB_STYLE = {
    "LG Chat": {"bg": "#1565C0", "fg": "#ffffff"},
    "NL Chat": {"bg": "#2E7D32", "fg": "#ffffff"},
}
CONNECT_STATE_STYLE = {
    "AVAILABLE":    ("#E3F2FD", "#1565C0"),
    "READY":        ("#E3F2FD", "#1565C0"),
    "BREAK":        ("#FFE0B2", "#BF360C"),
    "LUNCH":        ("#C8E6C9", "#1B5E20"),
    "COACHING":     ("#EDE7F6", "#4A148C"),
    "TRAINING":     ("#E8EAF6", "#283593"),
    "TEAM MEETING": ("#E8EAF6", "#283593"),
    "OFFLINEWORK":  ("#FFCCBC", "#BF360C"),
    "NOT READY":    ("#FFEBEE", "#C62828"),
    "NOTREADY":     ("#FFEBEE", "#C62828"),
    "UNAVAILABLE":  ("#FFEBEE", "#C62828"),
    "ENDOFSHIFT":   ("#FAFAFA", "#757575"),
    "LOGIN":        ("#FFF8E1", "#F57F17"),
    "PERSONAL":     ("#FBE9E7", "#BF360C"),
    "DROPPED":      ("#B71C1C", "#ffffff"),
}
ATD_CODE_STYLE = {
    'PR':  {'bg': '#1B5E20', 'fg': '#ffffff'},
    'ABS': {'bg': '#B71C1C', 'fg': '#ffffff'},
    'AL':  {'bg': '#1A237E', 'fg': '#ffffff'},
    'LWP': {'bg': '#4A148C', 'fg': '#ffffff'},
    'CO':  {'bg': '#006064', 'fg': '#ffffff'},
    'SL':  {'bg': '#37474F', 'fg': '#ffffff'},
    'EL':  {'bg': '#37474F', 'fg': '#ffffff'},
    'ML':  {'bg': '#37474F', 'fg': '#ffffff'},
}
EXCLUDE_ORACLE_IDS = {
    '102884675', '103177016', '103125229', '102509109', '102855952',
    '102029874', '102371964', '102478278', '103117188', '103132423',
    '103085298', '103110013', '102482237', '102456984', '103002947',
    '102477293', '102369579', '102173363', '102458990', '103109343',
    '103060459',
}

# %%
def convert_to_datetime(st):
    return datetime(*st[:6])

def input_data_all_raw(data_dir):
    files=[]
    for fn in pathlib.Path(data_dir).glob('**/*.*'):
        if fn.name.startswith('_'): continue
        sfx=fn.suffixes
        if not (sfx and sfx[-1].lower() in ['.xlsx','.csv']): continue
        try:
            exp_dt=convert_to_datetime(time.localtime(os.path.getmtime(fn)))
            if sfx[-1].lower()=='.xlsx': df=pl.read_excel(fn)
            else:
                if os.path.getsize(fn)==0:
                    try: fn.unlink()
                    except: pass
                    continue
                df=pl.read_csv(fn,infer_schema_length=10000)
            if df.is_empty(): continue
            df=df.with_columns(pl.lit(exp_dt).alias('Export time'))
            files.append(df)
        except: continue
    return pl.concat(files,how='diagonal_relaxed') if files else pl.DataFrame()

def parse_shift_start(shift_str):
    try:
        start=str(shift_str).strip().split('-')[0]
        if len(start)==4 and start.isdigit(): return dtime(int(start[:2]),int(start[2:]))
    except: pass
    return None

def is_overnight(shift_str):
    try:
        p=str(shift_str).strip().split('-')
        if len(p)==2 and all(len(x)==4 and x.isdigit() for x in p): return int(p[1][:2])<=8
    except: pass
    return False

def should_include(row, today, yesterday, now_dt):
    shift=str(row['Shift']).strip(); upper=shift.upper()
    if upper in EXCLUDE_SHIFTS or 'TERMINAT' in upper: return False
    if upper in LEAVE_CODES or '-' not in shift: return row['Sched_Date']==today
    start_dt=row['Shift_Start_DT']
    if start_dt is None: return False
    if row['Sched_Date']==today: return start_dt<=now_dt
    elif row['Sched_Date']==yesterday:
        if not is_overnight(shift): return False
        try:
            e=shift.strip().split('-')[1]
            end_dt=datetime.combine(today,dtime(int(e[:2]),int(e[2:])))
            return end_dt>=now_dt
        except: return False
    return False

def calc_attendance(row):
    shift=str(row['Shift']).strip(); upper=shift.upper()
    if upper in LEAVE_CODES or '-' not in shift: return upper, None
    login_dt=row.get('Login_DT'); start_dt=row['Shift_Start_DT']
    if pd.isna(login_dt) or start_dt is None: return 'ABS', None
    diff=(login_dt-start_dt).total_seconds()
    if diff<-12*3600: diff+=86400
    elif diff>12*3600: diff-=86400
    if diff>60:
        h,rem=divmod(int(diff),3600); m,s=divmod(rem,60)
        return 'PR', f'{h:02d}:{m:02d}:{s:02d}'
    return 'PR', None

def build_atd_summary(df, group_col, sort_key_fn=None, grand_total=True):
    groups = df[group_col].unique().tolist()
    groups = sorted(groups, key=sort_key_fn) if sort_key_fn else sorted(groups)
    rows = []
    for g in groups:
        gdf       = df[df[group_col] == g]
        scheduled = len(gdf)
        present   = int((gdf["ATD Code"] == "PR").sum())
        planned   = int(gdf["ATD Code"].isin(list(LEAVE_CODES)).sum())
        abs_unp   = int((gdf["ATD Code"] == "ABS").sum())
        working   = present + abs_unp
        rows.append({
            group_col:     g,
            "Scheduled":   scheduled,
            "Present":     present,
            "Planned":     planned,
            "ABS":         abs_unp,
            "% Attend":    f"{round(present/working*100,1)}%" if working>0 else "N/A",
            "% Shrinkage": f"{round(abs_unp/working*100,1)}%" if working>0 else "N/A",
        })
    result = pd.DataFrame(rows)
    if grand_total and not result.empty:
        t_scheduled=int(result["Scheduled"].sum()); t_present=int(result["Present"].sum())
        t_planned=int(result["Planned"].sum());     t_abs=int(result["ABS"].sum())
        t_working=t_present+t_abs
        result=pd.concat([result,pd.DataFrame([{
            group_col:"Grand Total","Scheduled":t_scheduled,"Present":t_present,
            "Planned":t_planned,"ABS":t_abs,
            "% Attend":f"{round(t_present/t_working*100,1)}%" if t_working>0 else "N/A",
            "% Shrinkage":f"{round(t_abs/t_working*100,1)}%" if t_working>0 else "N/A",
        }])],ignore_index=True)
    return result

def get_shift_style(shift_val):
    t=parse_shift_start(str(shift_val))
    if t is None: return None, None
    h=t.hour
    if 5<=h<=10:    return "#2E7D32","#ffffff"
    elif 11<=h<=17: return "#E65100","#ffffff"
    else:           return "#1565C0","#ffffff"

def seconds_to_hms(s):
    try:
        s=int(float(s)); h,r=divmod(s,3600); m,sec=divmod(r,60)
        return f'{h:02d}:{m:02d}:{sec:02d}'
    except: return '—'

def hms_to_seconds(hms):
    try:
        p=str(hms).split(':'); return int(p[0])*3600+int(p[1])*60+int(p[2])
    except: return 0

def make_bar_cell(s,max_s,color):
    f=int((s/max_s if max_s>0 else 0)*150); e=150-f
    bg=color if color else "#43a047"
    return (
        '<td style="padding:3px 8px;border:1px solid #ddd;vertical-align:middle;" nowrap>'
        '<table cellpadding="0" cellspacing="0" border="0" width="150" style="border-collapse:collapse;"><tr height="14">'
        f'<td width="{f}" height="14" bgcolor="{bg}" style="height:14px;">{"&nbsp;"if f>0 else""}</td>'
        f'<td width="{e}" height="14" bgcolor="#e0e0e0" style="height:14px;">{"&nbsp;"if e>0 else""}</td>'
        '</tr></table></td>')

def build_html_table(df, title, is_global=False, cases=0, summary=''):
    now=datetime.now()-(timedelta(hours=14) if is_global else timedelta(0))
    tz='(PST)' if is_global else '(VNT)'
    sub=f"Updated {now.strftime('%d-%b-%Y')} · {now.strftime('%I:%M %p')} {tz}"
    df=df.copy().drop(columns=[c for c in ['sheet_name','Export time'] if c in df.columns])
    if 'Duration (s)' in df.columns:
        df['Duration']=df['Duration (s)'].apply(seconds_to_hms); df=df.drop(columns=['Duration (s)'])
    cols=list(df.columns)
    dur_idx     = cols.index('Duration')      if 'Duration'      in cols else -1
    lob_idx     = cols.index('LOB')           if 'LOB'           in cols else -1
    state_idx   = cols.index('Connect State') if 'Connect State' in cols else -1
    loc_idx     = cols.index('Location')      if 'Location'      in cols else -1
    shift_idx   = cols.index('Shift')         if 'Shift'         in cols else -1
    atd_idx     = cols.index('ATD Code')      if 'ATD Code'      in cols else -1
    late_idx    = cols.index('Late')          if 'Late'          in cols else -1
    pct_att_idx = cols.index('% Attend')      if '% Attend'      in cols else -1
    pct_shr_idx = cols.index('% Shrinkage')   if '% Shrinkage'   in cols else -1
    dur_secs=[hms_to_seconds(str(row[cols[dur_idx]])) for _,row in df.iterrows()] if dur_idx>=0 else [0]*len(df)
    max_s=max(dur_secs) if dur_secs and dur_idx>=0 else 1
    th='bgcolor="#1e3a5f" style="color:#ffffff;padding:7px 12px;border:1px solid #2c4f7c;text-align:left;" nowrap'
    hdrs=[f'<th {th}>{c}</th>' for c in cols]
    if dur_idx>=0: hdrs.insert(dur_idx+1,f'<th {th}>Duration Bar</th>')
    def tdp(bg,v): return f'<td bgcolor="{bg}" style="color:#1a1a1a;padding:6px 12px;border:1px solid #ddd;" nowrap>{v}</td>'
    def tdb(bg,fg,v,al='left'): return f'<td bgcolor="{bg}" style="color:{fg};padding:6px 12px;border:1px solid #ddd;font-weight:bold;text-align:{al};" nowrap>{v}</td>'
    rows_html=''
    for i,(_,row) in enumerate(df.iterrows()):
        rbg='#f0f4ff' if i%2==0 else '#ffffff'
        lv=str(row[cols[lob_idx]]) if lob_idx>=0 else ''
        cells=[]
        for j,col in enumerate(cols):
            v=str(row[col]) if pd.notna(row[col]) else '—'
            if j==dur_idx:
                try: s=hms_to_seconds(v); bg='#b71c1c' if s>3600 else '#e53935' if s>1800 else '#fb8c00' if s>900 else '#fdd835' if s>300 else None; fg='#ffffff' if bg and bg!='#fdd835' else '#1a1a1a' if bg else None
                except: bg=fg=None
                cells.append(tdb(bg,fg,v) if bg else tdp(rbg,v))
                cells.append(make_bar_cell(dur_secs[i],max_s,bg))
            elif j==loc_idx:
                s=LOC_STYLE.get(v); cells.append(tdb(s['bg'],s['fg'],v,'center') if s else tdp(rbg,v))
            elif j==lob_idx:
                s=LOB_STYLE.get(lv); cells.append(tdb(s['bg'],s['fg'],v,'center') if s else tdp(rbg,v))
            elif j==state_idx:
                sc=CONNECT_STATE_STYLE.get(v); cells.append(tdb(sc[0],sc[1],v) if sc else tdp(rbg,v))
            elif j==shift_idx:
                sbg,sfg=get_shift_style(v); cells.append(tdb(sbg,sfg,v,'center') if sbg else tdp(rbg,v))
            elif j==atd_idx:
                s=ATD_CODE_STYLE.get(v,{'bg':rbg,'fg':'#1a1a1a'}); cells.append(tdb(s['bg'],s['fg'],v,'center'))
            elif j==late_idx:
                cells.append(tdb('#E65100','#ffffff',v) if v not in ('—','None','nan','') else tdp(rbg,'—'))
            elif j==pct_att_idx:
                try:
                    n=float(v.replace('%',''))
                    bg='#1B5E20' if n>=90 else '#F57F17' if n>=75 else '#B71C1C'
                    cells.append(tdb(bg,'#ffffff',v,'center'))
                except: cells.append(tdp(rbg,v))
            elif j==pct_shr_idx:
                try:
                    n=float(v.replace('%',''))
                    bg='#B71C1C' if n>=20 else '#F57F17' if n>=10 else '#1B5E20'
                    cells.append(tdb(bg,'#ffffff',v,'center'))
                except: cells.append(tdp(rbg,v))
            else:
                cells.append(tdp(rbg,v))
        rows_html+=f"<tr>{''.join(cells)}</tr>"
    sh=f'  <span style="font-size:11px;">📊 {summary}</span><br>\n' if summary else ''
    return (f'<p>\n  <b style="color:#c0392b;font-size:16px;">🔴 {title}</b><br>\n'
            f'  <span style="font-size:12px;">{sub} &nbsp;|&nbsp; ⚡ <b>{cases} CASES</b></span><br>\n{sh}</p>\n'
            f'<div style="overflow-x:auto;">\n<table border="1" cellpadding="0" cellspacing="0" '
            f'style="border-collapse:collapse;font-size:12px;font-family:Segoe UI,Arial,sans-serif;">\n'
            f'  <thead><tr>{"".join(hdrs)}</tr></thead>\n  <tbody>{rows_html}</tbody>\n</table>\n</div>')

def send_html_via_webhook(df, title, is_global=False, cases=0, summary=''):
    if df is None or (hasattr(df,'empty') and df.empty):
        print(f"⏭️  Skipping '{title}'"); return
    payload={'html':build_html_table(df,title,is_global,cases,summary)}
    try:
        r=requests.post(TEAMS_WEBHOOK_URL,headers={'Content-Type':'application/json'},
                        data=json.dumps(payload),timeout=30)
        print(f"✅ Sent: '{title}'" if r.status_code in (200,202) else f"❌ Failed [{r.status_code}]: {r.text[:200]}")
    except Exception as e:
        print(f"❌ {e}")

# %%
# ── Time context ──────────────────────────────────────────────
now_dt    = datetime.now()
today     = (now_dt-timedelta(days=1)).date() if now_dt.hour<5 else now_dt.date()
yesterday = today-timedelta(days=1)

# ── Load all snapshots ────────────────────────────────────────
raw_all=input_data_all_raw(DATA_DIR)
if raw_all.is_empty(): raise RuntimeError("❌ No data")
print(f"✅ All snapshots: {len(raw_all)} rows")

lob_expr=pl.lit(None).cast(pl.Utf8)
for lbl,qs in LOB_MAP.items():
    lob_expr=pl.when(pl.col("Queue Group / Routing Profile").is_in(qs)).then(pl.lit(lbl)).otherwise(lob_expr)

raw_hcm=(raw_all
    .with_columns(lob_expr.alias("LOB"))
    .filter(pl.col("LOB").is_in(["NL Chat","LG Chat"]))
    .with_columns(
        pl.when(pl.col("Business Location").str.contains("Ho Chi Minh")).then(pl.lit("HCM"))
        .when(pl.col("Business Location").str.contains("Pune")).then(pl.lit("PUN"))
        .when(pl.col("Business Location").str.contains("Kolkata")).then(pl.lit("KOL"))
        .when(pl.col("Business Location").str.contains("Cairo")).then(pl.lit("CAI"))
        .otherwise(pl.lit("OTHER")).alias("Location"))
    .filter(pl.col("Location")=="HCM"))

email_lob_map=(raw_hcm.select(["Agent Email","LOB"]).drop_nulls()
    .unique(subset=["Agent Email"],keep="last").to_pandas()
    .assign(Email_key=lambda df:df["Agent Email"].str.strip().str.lower())
    .set_index("Email_key")["LOB"].to_dict())

HCM_CACHE_FILE = os.path.join(DATA_DIR, "_hcm_email_cache.csv")
if os.path.exists(HCM_CACHE_FILE):
    cache_df  = pd.read_csv(HCM_CACHE_FILE)
    cache_lob = (cache_df.dropna(subset=["LOB"])
                 .drop_duplicates("Email_key",keep="last")
                 .set_index("Email_key")["LOB"].to_dict())
else:
    cache_df  = pd.DataFrame(columns=["Email_key","LOB"])
    cache_lob = {}

session_hcm = (
    raw_all.filter(pl.col("Business Location").str.contains("Ho Chi Minh"))
    .select(["Agent Email"]).drop_nulls().unique().to_pandas()
    .assign(Email_key=lambda df: df["Agent Email"].str.strip().str.lower())[["Email_key"]]
)
session_hcm["LOB"] = session_hcm["Email_key"].map(email_lob_map)
updated_cache = (pd.concat([cache_df, session_hcm])
    .sort_values("LOB",na_position="last")
    .drop_duplicates("Email_key",keep="first")
    .reset_index(drop=True))
updated_cache.to_csv(HCM_CACHE_FILE, index=False)
email_lob_map  = {**cache_lob, **email_lob_map}
hcm_email_keys = set(updated_cache["Email_key"].tolist())
print(f"✅ Email→LOB: {len(email_lob_map)} | HCM cache: {len(hcm_email_keys)} agents")

# ── Build login_today: Login Time from latest snapshot per agent
lr = (raw_hcm.select(["Agent Name","Agent Email","Login Time","Export time"])
    .to_pandas().dropna(subset=["Agent Email","Login Time"]))
lr["Login_DT"]   = pd.to_datetime(lr["Login Time"], errors="coerce")
lr["Login_Date"] = lr["Login_DT"].dt.date
lr["Email_key"]  = lr["Agent Email"].str.strip().str.lower()
login_today = (lr[lr["Login_Date"] == today].dropna(subset=["Login_DT"])
    .sort_values("Export time", ascending=False)
    .drop_duplicates("Email_key", keep="first")
    [["Email_key","Agent Name","Agent Email","Login_DT"]].reset_index(drop=True))
print(f"✅ Logins today: {len(login_today)}")

# ── Latest Connect State ──────────────────────────────────────
state_latest=(raw_hcm.filter(pl.col("Export time")==pl.col("Export time").max())
    .select(["Agent Email","Connect State"]).to_pandas()
    .assign(Email_key=lambda df:df["Agent Email"].str.strip().str.lower())
    .drop_duplicates("Email_key",keep="last")[["Email_key","Connect State"]].reset_index(drop=True))

# ── Read Schedule ─────────────────────────────────────────────
sched_raw = pd.read_excel(SCHEDULE_FILE, dtype=str)
sched_raw.columns = [str(c).strip() for c in sched_raw.columns]
date_cols_raw = [c for c in sched_raw.columns if re.match(r'^\d{4}-\d{2}-\d{2}', c)]
col_map       = {c: c[:10] for c in date_cols_raw}
sched_raw     = sched_raw.rename(columns=col_map)

oracle_col = 'OracleID' if 'OracleID' in sched_raw.columns else None
if oracle_col:
    before    = len(sched_raw)
    sched_raw = sched_raw[~sched_raw[oracle_col].astype(str).str.strip().isin(EXCLUDE_ORACLE_IDS)].reset_index(drop=True)
    print(f"✅ Excluded {before-len(sched_raw)} OracleIDs | Remaining: {len(sched_raw)}")
else:
    print("⚠️ OracleID column not found — skipping exclusion")

keep_dates=[str(d) for d in [yesterday,today] if str(d) in col_map.values()]
print(f"✅ Schedule dates: {keep_dates}")

email_col="Email"; name_col="Employee Name" if "Employee Name" in sched_raw.columns else None
id_cols=[email_col]+([name_col] if name_col else [])
sched_long=(sched_raw[id_cols+keep_dates]
    .melt(id_vars=id_cols,value_vars=keep_dates,var_name="Sched_Date",value_name="Shift")
    .assign(
        Sched_Date=lambda df:pd.to_datetime(df["Sched_Date"]).dt.date,
        Email_key=lambda df:df[email_col].astype(str).str.strip().str.lower(),
        Shift=lambda df:df["Shift"].astype(str).str.strip())
    .loc[lambda df:~df["Shift"].isin(["","nan","None"])]
    .drop_duplicates(["Email_key","Sched_Date"]).reset_index(drop=True))

today_sched = sched_long[sched_long["Sched_Date"]==today].copy()
today_sched["_su"] = today_sched["Shift"].str.strip().str.upper()
valid_emails = today_sched[
    (today_sched["_su"]!="") & (today_sched["_su"]!="NAN") &
    (~today_sched["_su"].isin(EXCLUDE_SHIFTS)) &
    (~today_sched["_su"].str.contains("TERMINAT",na=False))
]["Email_key"].unique()
sched_long = sched_long[sched_long["Email_key"].isin(valid_emails)].reset_index(drop=True)
print(f"✅ Valid agents for {today}: {len(valid_emails)} | sched_long: {len(sched_long)} rows")

sched_long["Shift_Start_DT"]=sched_long.apply(
    lambda r:(datetime.combine(r["Sched_Date"],parse_shift_start(r["Shift"]))
              if "-" in str(r["Shift"]) and parse_shift_start(r["Shift"]) else None),axis=1)

sched_active=(sched_long[sched_long.apply(lambda r:should_include(r,today,yesterday,now_dt),axis=1)]
    .copy().assign(_sort_key=lambda df:df["Shift_Start_DT"].fillna(pd.Timestamp("2099-01-01")))
    .sort_values("_sort_key").reset_index(drop=True))
print(f"✅ Active shifts: {len(sched_active)}")

# ── Merge + ATD ───────────────────────────────────────────────
merged=sched_active.merge(login_today,on="Email_key",how="left")
merged=merged.merge(state_latest,on="Email_key",how="left")
if name_col: merged["Agent Name"]=merged["Agent Name"].fillna(merged[name_col])
merged["Agent Email"]=merged["Agent Email"].fillna(merged[email_col])
merged["LOB"]=merged["Email_key"].map(email_lob_map).fillna("—")
merged["Connect State"]=merged.apply(
    lambda r:"DROPPED" if (pd.notna(r.get("Login_DT")) and
        (pd.isna(r.get("Connect State")) or str(r.get("Connect State")).strip() in ["","nan","None"]))
    else (r.get("Connect State") if pd.notna(r.get("Connect State")) else "—"),axis=1)

res=merged.apply(calc_attendance,axis=1,result_type="expand")
merged["ATD Code"]=res[0]; merged["Late"]=res[1]
merged["Login Time"]=merged["Login_DT"].apply(lambda x:x.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(x) else "—")

attendance_pd=(merged[["Agent Name","Agent Email","Login Time","Shift","LOB","Connect State","Late","ATD Code"]]
    .reset_index(drop=True))
_ls=attendance_pd["LOB"].astype(str).str.strip()
_valid=_ls.notna()&(_ls!="")&(_ls!="nan")&(_ls!="None")&(_ls!="—")&(_ls!="-")
_is_leave=attendance_pd["ATD Code"].isin(list(LEAVE_CODES))
attendance_pd=attendance_pd[_valid|_is_leave].reset_index(drop=True)

pr_n=(attendance_pd["ATD Code"]=="PR").sum()
abs_n=(attendance_pd["ATD Code"]=="ABS").sum()
late_n=attendance_pd["Late"].notna().sum()
lv_n=(~attendance_pd["ATD Code"].isin(["PR","ABS"])).sum()
print(f"\n{'─'*50}\nAttendance: {len(attendance_pd)} | PR:{pr_n} | ABS:{abs_n} | Late:{late_n} | Leave:{lv_n}\n{'─'*50}")

# ── ATD Summary ───────────────────────────────────────────────
lob_summary_pd=build_atd_summary(attendance_pd,"LOB")
attendance_pd["_shift_group"]=attendance_pd["Shift"].apply(
    lambda s:"Planned" if str(s).upper() in LEAVE_CODES else s)

def _shift_sort_key(g):
    if g=="Planned": return "9999"
    t=parse_shift_start(str(g))
    return f"{t.hour:02d}{t.minute:02d}" if t else "8888"

shift_summary_pd=(build_atd_summary(attendance_pd,"_shift_group",sort_key_fn=_shift_sort_key)
    .rename(columns={"_shift_group":"Shift"}))

total_day_hc=len(sched_long[
    (sched_long["Sched_Date"]==today)&
    (~sched_long["Shift"].str.upper().isin(EXCLUDE_SHIFTS))&
    (~sched_long["Shift"].str.upper().str.contains("TERMINAT"))])
total_planned=int(attendance_pd["ATD Code"].isin(list(LEAVE_CODES)).sum())
if "Planned" in shift_summary_pd["Shift"].values:
    shift_summary_pd.loc[shift_summary_pd["Shift"]=="Planned","% Shrinkage"]=(
        f"{round(total_planned/total_day_hc*100,1)}%" if total_day_hc>0 else "N/A")
print(f"Total day HC: {total_day_hc} | Planned: {total_planned}")

# ── Send ──────────────────────────────────────────────────────
atd_total=len(attendance_pd)
atd_summary=f"PR: {pr_n} | ABS: {abs_n} | Late: {late_n} | Leave: {lv_n}"
send_html_via_webhook(lob_summary_pd,   "ATD Summary by LOB",   cases=atd_total, summary=atd_summary)
send_html_via_webhook(shift_summary_pd, "ATD Summary by Shift",  cases=atd_total)

shift_order=sorted(attendance_pd["_shift_group"].unique().tolist(),key=_shift_sort_key)
for sv in shift_order:
    sdf=(attendance_pd[attendance_pd["_shift_group"]==sv].drop(columns=["_shift_group"]).reset_index(drop=True))
    if sdf.empty: continue
    _pr=(sdf["ATD Code"]=="PR").sum(); _abs=(sdf["ATD Code"]=="ABS").sum(); _late=sdf["Late"].notna().sum()
    label="Planned" if sv=="Planned" else f"Shift {sv}"
    send_html_via_webhook(sdf,f"Attendance — VN | {label}",cases=len(sdf),summary=f"PR: {_pr} | ABS: {_abs} | Late: {_late}")


