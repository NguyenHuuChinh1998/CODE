# %%
import os, io, time, pathlib, json, requests
import pandas as pd
import polars as pl
from datetime import datetime, timedelta

TEAMS_WEBHOOK_URL = "https://default599e51d62f8c43478e591f795a51a9.8c.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/8a7e8a2d23674e2797e45eb2d9b42941/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=w1-Grzliw76VIaSnhS2VkBrnFzl2tdev6Brm1KdeWeM"
DATA_DIR = r"C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\Rawdata\CAPTURE\lc_rawdata_in_console"

LOB_MAP = {
    "LG Chat":  ["Chat_OD_EN_Car_Activity","Chat_OD_EN_Lodging","Chat - Global English Lodging Nesting","Chat_Lodging English w Car","Chat_AC_GLB_EN_Lodging_Proficient","Chat_AC_GLB_EN_Car_Activity","Chat_AC_GLB_EN_Lodging_Expert"],
    "NL Chat":  ["Chat - Global English Non- Lodging Nesting","Chat_OD_EN_Dual_GDS","Chat_AC_GLB_EN_Proficient","Chat_AC_GLB_EN_Expert"],
    "NL Voice": ["Voice_OD_Proficient_GLB_EN","Voice_OD_Expert_GLB_EN","Voice_AC_GLB_EN_Proficient"],
    "LG Voice": ["Voice_OD_GLB_EN_Lodging_Proficient","Voice_OD_GLB_EN_Lodging_Expert","Voice_AC_GLB_EN_Lodging_Proficient","Voice_AC_GLB_EN_Lodging_Nesting"],
}
LOC_STYLE = {
    "HCM": {"bg":"#DA251D","fg":"#FFD700"}, "VN":  {"bg":"#DA251D","fg":"#FFD700"},
    "PUN": {"bg":"#388E3C","fg":"#ffffff"},  "KOL": {"bg":"#1565C0","fg":"#ffffff"},
    "CAI": {"bg":"#E65100","fg":"#ffffff"},
}
LOB_STYLE = {
    "LG Chat":  {"bg":"#1565C0","fg":"#ffffff"}, "NL Chat":  {"bg":"#2E7D32","fg":"#ffffff"},
    "LG Voice": {"bg":"#4A148C","fg":"#ffffff"}, "NL Voice": {"bg":"#BF360C","fg":"#ffffff"},
}

def send_banner():
    ts   = datetime.now().strftime("%d-%b-%Y  %I:%M %p (VNT)")
    html = (
        '<table cellpadding="0" cellspacing="0" border="0" '
        'style="border-collapse:collapse;width:100%;border-left:5px solid #E65100;">'
        '<tr>'
        '<td width="5" bgcolor="#E65100" style="width:5px;">&nbsp;</td>'
        '<td style="padding:8px 14px;">'
        '<span style="font-size:18px;">🚩</span>&nbsp;'
        '<b style="color:#E65100;font-size:16px;">Long Chat Flag Report</b><br>'
        '<span style="font-size:11px;opacity:0.75;">'
        f'⏱ <b>{ts}</b>&nbsp;&nbsp;|&nbsp;&nbsp;'
        'Agents with long-running chat sessions'
        '</span></td></tr></table>'
    )
    try:
        r = requests.post(TEAMS_WEBHOOK_URL, headers={"Content-Type": "application/json"},
                          data=json.dumps({"html": html}), timeout=30)
        print(f"[BANNER] Sent" if r.status_code in (200, 202) else f"[BANNER] Failed [{r.status_code}]")
    except Exception as e:
        print(f"[BANNER] Error: {e}")

# %%
def convert_to_datetime(st):
    return datetime(*st[:6])

def input_data(data_dir):
    list_files=[]
    for filename in pathlib.Path(data_dir).glob('**/*.*'):
        sfx=filename.suffixes
        if not (sfx and sfx[-1].lower() in ['.xlsx','.csv']): continue
        exp_dt=convert_to_datetime(time.localtime(os.path.getmtime(filename)))
        try:
            if sfx[-1].lower()=='.xlsx':
                df=pl.read_excel(filename)
            else:
                if os.path.getsize(filename)==0: continue
                df=pl.read_csv(filename,infer_schema_length=10000)
            if df.is_empty(): continue
            df=df.with_columns([pl.lit(filename.stem).alias('sheet_name'),pl.lit(exp_dt).alias('Export time')])
            list_files.append(df)
        except Exception as e: print(f"❌ Error reading {filename.name}: {e}")
    return pl.concat(list_files,how='diagonal_relaxed') if list_files else pl.DataFrame()

def str_hms_to_seconds(hms):
    try:
        parts=[int(p) for p in str(hms).split(':')]
        if len(parts)==3: return parts[0]*3600+parts[1]*60+parts[2]
        if len(parts)==2: return parts[0]*60+parts[1]
        return int(parts[0])
    except: return None

def process_lc(df):
    df=df.with_columns(
        pl.when((pl.col("LOB").is_in(["LG Chat","LG Voice"]))&(pl.col("Duration (s)")>=900)).then(True)
        .when((pl.col("LOB").is_in(["NL Chat","NL Voice"]))&(pl.col("Duration (s)")>=1500)).then(True)
        .otherwise(False).alias("LC"))
    df=df.filter(pl.col("LC")==True).sort("Duration (s)",descending=True)
    LC_Cases=df.shape[0]
    return df.head(25).select(["Location","Agent Name","Manager Name","Connected Time","Duration (s)","LOB"]), LC_Cases

def seconds_to_hms(s):
    try:
        s=int(float(s)); h,r=divmod(s,3600); m,sec=divmod(r,60)
        return f'{h:02d}:{m:02d}:{sec:02d}'
    except: return '—'

def hms_to_seconds(hms):
    try:
        p=str(hms).split(':'); return int(p[0])*3600+int(p[1])*60+int(p[2])
    except: return 0

def get_duration_color(hms_str, lob):
    try:
        s=hms_to_seconds(hms_str); target=900 if "LG" in str(lob) else 1500
        ratio=s/target if target>0 else 0; pct=(ratio-1)*100
        if pct>90: return "#b71c1c","#ffffff"
        if pct>60: return "#e53935","#ffffff"
        if pct>30: return "#fb8c00","#ffffff"
        if pct>0:  return "#43a047","#ffffff"
    except: pass
    return None, None

def make_bar_cell(s, max_s, color):
    ratio  = s / max_s if max_s > 0 else 0
    filled = int(ratio * 15)
    empty  = 15 - filled
    text_color = color if color else "#43a047"
    bar    = '█' * filled + '░' * empty
    return (
        f'<td style="padding:3px 8px;border:1px solid #ddd;'
        f'font-family:monospace;letter-spacing:1px;'
        f'color:{text_color};font-weight:bold;" nowrap>{bar}</td>'
    )

def build_html_table(df, title, is_global=False, lc_cases=0):
    now=datetime.now()-(timedelta(hours=14) if is_global else timedelta(0))
    tz='(PST)' if is_global else '(VNT)'
    subtitle=f"Updated {now.strftime('%d-%b-%Y')} · {now.strftime('%I:%M %p')} {tz}"
    df=df.copy().drop(columns=[c for c in ['Duration Bar','sheet_name','Export time'] if c in df.columns])
    if 'Duration (s)' in df.columns:
        df['Duration']=df['Duration (s)'].apply(seconds_to_hms); df=df.drop(columns=['Duration (s)'])
    col_names=list(df.columns)
    dur_idx=col_names.index('Duration') if 'Duration' in col_names else -1
    lob_idx=col_names.index('LOB')      if 'LOB'      in col_names else -1
    loc_idx=col_names.index('Location') if 'Location' in col_names else -1
    dur_secs=[hms_to_seconds(str(row[col_names[dur_idx]])) for _,row in df.iterrows()] if dur_idx>=0 else [0]*len(df)
    max_s=max(dur_secs) if dur_secs else 1
    th='bgcolor="#1e3a5f" style="color:#ffffff;padding:7px 12px;border:1px solid #2c4f7c;text-align:left;" nowrap'
    header_list=[f'<th {th}>{c}</th>' for c in col_names]
    if dur_idx>=0: header_list.insert(dur_idx+1,f'<th {th}>Duration Bar</th>')
    rows_html=""
    for i,(_,row) in enumerate(df.iterrows()):
        cells=[]; lob_val=str(row[col_names[lob_idx]]) if lob_idx>=0 else ""
        for j,col in enumerate(col_names):
            val=str(row[col]) if pd.notna(row[col]) else "—"
            if j==loc_idx:
                s=LOC_STYLE.get(val)
                cells.append(f'<td bgcolor="{s["bg"]}" style="color:{s["fg"]};padding:6px 12px;border:1px solid #ddd;font-weight:bold;" nowrap>{val}</td>' if s
                              else f'<td style="padding:6px 12px;border:1px solid #ddd;" nowrap>{val}</td>')
            elif j==dur_idx:
                bg,fg=get_duration_color(val,lob_val)
                cells.append(f'<td bgcolor="{bg}" style="color:{fg};padding:6px 12px;border:1px solid #ddd;font-weight:bold;" nowrap>{val}</td>' if bg
                              else f'<td style="padding:6px 12px;border:1px solid #ddd;" nowrap>{val}</td>')
                cells.append(make_bar_cell(dur_secs[i],max_s,bg))
            elif j==lob_idx:
                s=LOB_STYLE.get(lob_val)
                cells.append(f'<td bgcolor="{s["bg"]}" style="color:{s["fg"]};padding:6px 12px;border:1px solid #ddd;font-weight:bold;text-align:center;" nowrap>{val}</td>' if s
                              else f'<td style="padding:6px 12px;border:1px solid #ddd;" nowrap>{val}</td>')
            else:
                cells.append(f'<td style="padding:6px 12px;border:1px solid #ddd;" nowrap>{val}</td>')
        rows_html+=f"<tr>{''.join(cells)}</tr>"
    return (f'<p>\n  <b style="color:#c0392b;font-size:16px;">🔴 {title}</b><br>\n'
            f'  <span style="font-size:12px;">{subtitle} &nbsp;|&nbsp; ⚡ <b>{lc_cases} CASES</b></span>\n</p>\n'
            f'<div style="overflow-x:auto;">\n<table border="1" cellpadding="0" cellspacing="0" '
            f'style="border-collapse:collapse;font-size:12px;font-family:Segoe UI,Arial,sans-serif;">\n'
            f'  <thead><tr>{"".join(header_list)}</tr></thead>\n  <tbody>{rows_html}</tbody>\n</table>\n</div>')

def send_html_via_webhook(df, title, is_global=False, lc_cases=0):
    if lc_cases==0: print(f"⏭️  Skipping '{title}' — 0 cases"); return
    if df is None or df.empty: print(f"⏭️  Skipping '{title}' — empty"); return
    payload={"html":build_html_table(df,title,is_global,lc_cases)}
    try:
        resp=requests.post(TEAMS_WEBHOOK_URL,headers={"Content-Type":"application/json"},
                           data=json.dumps(payload),timeout=30)
        print(f"✅ Sent: '{title}' ({lc_cases} cases)" if resp.status_code in (200,202)
              else f"❌ Failed [{resp.status_code}]: {resp.text[:300]}")
    except Exception as e: print(f"❌ Error: {e}")

# ── Load & process ────────────────────────────────────────────
long_chat_db=input_data(DATA_DIR)
if long_chat_db.is_empty(): raise RuntimeError(f"❌ No data loaded from: {DATA_DIR}")
print(f"✅ Loaded {len(long_chat_db)} rows")

long_chat_db=long_chat_db.sort(["Export time"]).filter(pl.col("Export time")==pl.col("Export time").max())

lob_expr=pl.lit(None).cast(pl.Utf8)
for lob_label,queues in LOB_MAP.items():
    lob_expr=pl.when(pl.col("Queue Group / Routing Profile").is_in(queues)).then(pl.lit(lob_label)).otherwise(lob_expr)
long_chat_db=long_chat_db.with_columns(lob_expr.alias("LOB"))

long_chat_db=long_chat_db.with_columns(
    pl.when(pl.col("Business Location").str.contains("Ho Chi Minh")).then(pl.lit("VN"))
    .when(pl.col("Business Location").str.contains("Pune")).then(pl.lit("PUN"))
    .when(pl.col("Business Location").str.contains("Kolkata")).then(pl.lit("KOL"))
    .when(pl.col("Business Location").str.contains("Cairo")).then(pl.lit("CAI"))
    .otherwise(pl.lit("OTHER")).alias("Location"))

long_chat_db=long_chat_db.with_columns(
    pl.col("Connected Duration").cast(str).map_elements(str_hms_to_seconds,return_dtype=pl.Int64).alias("Duration (s)"))

hcm_lc       = long_chat_db.filter(pl.col("Location")=="VN")
global_lg_lc = long_chat_db.filter(pl.col("LOB")=="LG Chat")
global_nl_lc = long_chat_db.filter(pl.col("LOB")=="NL Chat")
global_lg_lv = long_chat_db.filter(pl.col("LOB")=="LG Voice")
global_nl_lv = long_chat_db.filter(pl.col("LOB")=="NL Voice")

hcm_lc_processed,    LC_Cases_hcm       = process_lc(hcm_lc)
global_lg_processed,  LC_Cases_global_lg = process_lc(global_lg_lc)
global_nl_processed,  LC_Cases_global_nl = process_lc(global_nl_lc)
global_lgv_processed, LV_Cases_global_lg = process_lc(global_lg_lv)
global_nlv_processed, LV_Cases_global_nl = process_lc(global_nl_lv)

hcm_lc_pd     = hcm_lc_processed.to_pandas()
global_lg_pd  = global_lg_processed.to_pandas()
global_nl_pd  = global_nl_processed.to_pandas()
global_lgv_pd = global_lgv_processed.to_pandas()
global_nlv_pd = global_nlv_processed.to_pandas()

print(f"VN: {LC_Cases_hcm} | LG Chat: {LC_Cases_global_lg} | NL Chat: {LC_Cases_global_nl} | LG Voice: {LV_Cases_global_lg} | NL Voice: {LV_Cases_global_nl}")

# %%
send_banner()

send_html_via_webhook(hcm_lc_pd,     "Long Chat Report for VN",                  is_global=False, lc_cases=LC_Cases_hcm)
# send_html_via_webhook(global_lg_pd,  "Lodging Long Chat Report for Global",      is_global=True,  lc_cases=LC_Cases_global_lg)
# send_html_via_webhook(global_nl_pd,  "Non-Lodging Long Chat Report for Global",  is_global=True,  lc_cases=LC_Cases_global_nl)
# send_html_via_webhook(global_lgv_pd, "Lodging Long Voice Report for Global",     is_global=True,  lc_cases=LV_Cases_global_lg)
# send_html_via_webhook(global_nlv_pd, "Non-Lodging Long Voice Report for Global", is_global=True,  lc_cases=LV_Cases_global_nl)


