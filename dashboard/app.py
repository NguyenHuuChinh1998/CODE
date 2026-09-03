"""
WFM Pipeline Control Dashboard — Expedia HCM
Run: streamlit run app.py
"""

import streamlit as st
import sqlite3
import subprocess
import threading
import json
import os
import re
import time
import tempfile
import shutil
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="WFM Pipeline — Expedia HCM",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# PATHS  (tự động detect theo user hiện tại)
# ─────────────────────────────────────────────────────────────
HOME     = Path(os.path.expanduser("~"))
CNX_BASE = HOME / "Concentrix Corporation" / "WFM-Expedia-HCM - Branding files"
CODE_DIR = CNX_BASE / "BI_Task" / "CODE" / "Python_Code"
RES_DIR  = CNX_BASE / "BI_Task" / "CODE" / "Resources"

DASH_DIR    = CODE_DIR / "dashboard"
DB_PATH     = DASH_DIR / "pipeline.db"
LOG_DIR     = DASH_DIR / "logs"
PREVIEW_DIR = DASH_DIR / "email_previews"

for d in [DASH_DIR, LOG_DIR, PREVIEW_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ─────────────────────────────────────────────────────────────
# TASK DEFINITIONS
# ─────────────────────────────────────────────────────────────
LAYERS = {
    0: ("📥", "Data Collection",  "#0D47A1"),
    1: ("⚙️",  "ETL Processing",  "#1B5E20"),
    2: ("🤖", "Realtime AUTOBOT", "#4A148C"),
    3: ("✉️", "Email Reports",    "#B71C1C"),
}

def nb(name):
    return str(CODE_DIR / name)

def py(name):
    return str(CODE_DIR / "AUTOBOT" / "py" / name)

TASKS = {
    # ── Layer 0: Data Collection ──────────────────────────────
    "download": {
        "name": "Download Raw Data",
        "type": "py", "script": py("download_test.py"),
        "deps": [], "layer": 0,
        "desc": "Selenium → Expedia Console → CAPTURE/",
    },
    "nice_iex": {
        "name": "NICE IEX Schedule",
        "type": "notebook", "notebook": nb("AUTOBOT/nice_download.ipynb"),
        "deps": [], "layer": 0,
        "desc": "NICE WFM → IEX schedule xlsx → INPUT_AGENT_IEX/",
    },

    # ── Layer 1: ETL Processing ───────────────────────────────
    "hc_master": {
        "name": "HC Master Pipeline",
        "type": "notebook", "notebook": nb("HC_Master_optimized__2_ (1).ipynb"),
        "deps": [], "layer": 1,
        "desc": "HRIS dump → hc_extend → hc_extend_combination.parquet",
        "output_check": str(RES_DIR / "hc_extend_combination.parquet"),
    },
    "iex": {
        "name": "IEX Processing",
        "type": "notebook", "notebook": nb("IEX_optimized.ipynb"),
        "deps": ["hc_master", "nice_iex"], "layer": 1,
        "desc": "IEX raw → clean → OUTPUT_AGENT_IEX_FOR_REPORT/",
    },
    "rta": {
        "name": "RTA Processing",
        "type": "notebook", "notebook": nb("RTA_optimized (1).ipynb"),
        "deps": ["iex", "download"], "layer": 1,
        "desc": "Agent Activity + IEX → STORAGE_OUTPUT_RTA/",
    },
    "atd": {
        "name": "ATD Final Pipeline",
        "type": "notebook", "notebook": nb("atd_optimized.ipynb"),
        "deps": ["rta", "hc_master"], "layer": 1,
        "desc": "Leave + Roster + RTA → ATD_Final.parquet",
        "output_check": str(RES_DIR / "ATD_Final.parquet"),
    },
    "performance": {
        "name": "Performance ETL",
        "type": "notebook", "notebook": nb("performance_3_aws_migration.ipynb"),
        "deps": ["hc_master"], "layer": 1,
        "desc": "AWS Connect schema → _performance_hcm.parquet",
        "output_check": str(RES_DIR / "_performance_hcm.parquet"),
    },
    "miv": {
        "name": "MIV Report",
        "type": "notebook", "notebook": nb("MIV_v2.ipynb"),
        "deps": ["performance"], "layer": 1,
        "desc": "Survey + T3 + AFCR → MIV_Data/",
    },
    "non_compliance": {
        "name": "Non-Compliance",
        "type": "notebook", "notebook": nb("Non Compliance.ipynb"),
        "deps": ["rta", "iex"], "layer": 1,
        "desc": "IEX intervals + RTA → Adherence + Rest compliance",
    },
    "ot": {
        "name": "OT Reconciliation",
        "type": "notebook", "notebook": nb("OT.ipynb"),
        "deps": ["rta"], "layer": 1,
        "desc": "Roster × OT Plan × RAMCO → OT_CONFIRMATION_REPORT.xlsx",
    },
    "global_hc": {
        "name": "Global HC Process",
        "type": "notebook", "notebook": nb("Global_HC_Process.ipynb"),
        "deps": [], "layer": 1,
        "desc": "17 locations Operations HC → long format",
    },
    "schedule_tool": {
        "name": "Schedule Tool",
        "type": "notebook", "notebook": nb("Schedule_Tool_optimized.ipynb"),
        "deps": ["hc_master"], "layer": 1,
        "desc": "Schedule generation + swap chain (networkx)",
    },

    # ── Layer 2: Realtime AUTOBOT ─────────────────────────────
    "bot_atd": {
        "name": "ATD Realtime Bot",
        "type": "py", "script": py("atd_realtime.py"),
        "deps": [], "layer": 2,
        "desc": "Agent snapshots → detect NCNS/absent → Teams",
    },
    "bot_ic": {
        "name": "IC Monitor Bot",
        "type": "py", "script": py("ic_monitor.py"),
        "deps": [], "layer": 2,
        "desc": "UCP + IEX intervals → IC compliance → Teams",
    },
    "bot_lc": {
        "name": "Long Chat Flag Bot",
        "type": "py", "script": py("lc_flag.py"),
        "deps": [], "layer": 2,
        "desc": "Live chat durations → flag long chats → Teams",
    },
    "bot_ou": {
        "name": "OU Present Bot",
        "type": "py", "script": py("ou_present.py"),
        "deps": [], "layer": 2,
        "desc": "Global OU files → Required Heads → Teams",
    },
    "bot_ic_summary": {
        "name": "IC Summary Bot",
        "type": "py", "script": py("ic_summary.py"),
        "deps": [], "layer": 2,
        "desc": "Forecast vs actual interval → Heads Contribution → Teams",
    },

    # ── Layer 3: Email Reports ────────────────────────────────
    "email_atd": {
        "name": "ATD Report",
        "type": "notebook", "notebook": nb("Send_Email/atd_auto_send_email.ipynb"),
        "deps": ["atd"], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "Attendance Mismatch & NM Report → Outlook",
    },
    "email_atd_rt": {
        "name": "ATD Realtime Email",
        "type": "notebook", "notebook": nb("Send_Email/atd_realtime_send_mail.ipynb"),
        "deps": [], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "Night shift ATD (2200-0600) → Teams + Outlook",
    },
    "email_adherence": {
        "name": "Adherence Report",
        "type": "notebook", "notebook": nb("Send_Email/adherence_auto_send_email.ipynb"),
        "deps": ["atd", "non_compliance"], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "Schedule Adherence MTD/weekly → Outlook",
    },
    "email_compliance": {
        "name": "Agent Compliance",
        "type": "notebook", "notebook": nb("Send_Email/agent_compliance_auto_send_email.ipynb"),
        "deps": ["atd", "non_compliance"], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "Adherence Gap + Rest Compliance → Outlook",
    },
    "email_missed": {
        "name": "Missed Productive",
        "type": "notebook", "notebook": nb("Send_Email/missed_productive_send_email.ipynb"),
        "deps": ["atd"], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "Missed/Late/Leave MoM + 7-day + D-1 → Outlook",
    },
    "email_performance": {
        "name": "Performance Report",
        "type": "notebook", "notebook": nb("Send_Email/performance_auto_send_email.ipynb"),
        "deps": ["performance"], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "KPI performance 1.2M rows → Outlook",
    },
    "email_ic": {
        "name": "IC Report",
        "type": "notebook", "notebook": nb("Send_Email/ic_report_auto_send_email.ipynb"),
        "deps": [], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "Interval Compliance MTD/weekly → Outlook",
    },
    "email_headcount": {
        "name": "Headcount Report",
        "type": "notebook", "notebook": nb("Send_Email/headcount_report_send_mail.ipynb"),
        "deps": ["hc_master"], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "HC + Attrition report → Outlook",
    },
    "email_team_align": {
        "name": "Team Alignment",
        "type": "notebook", "notebook": nb("Send_Email/team_alignment.ipynb"),
        "deps": ["hc_master"], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "Team Alignment weekly snapshot → Outlook",
    },
    "email_schedule": {
        "name": "Schedule Email",
        "type": "notebook", "notebook": nb("Send_Email/schedule_send_mail.ipynb"),
        "deps": [], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "Draft + Final schedule email → Outlook Reply-All",
    },
    "email_deactivate": {
        "name": "Okta Deactivation",
        "type": "notebook", "notebook": nb("Send_Email/deactivate_okta_send_mail.ipynb"),
        "deps": ["hc_master"], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "Okta deactivation request → CNX.EXPEDIA.IDAdmins",
    },
    "email_no_rta": {
        "name": "No RTA Coverage",
        "type": "notebook", "notebook": nb("Send_Email/no_rta_coverage.ipynb"),
        "deps": [], "layer": 3,
        "send_var": "SEND_EMAIL",
        "desc": "No coverage period → WFM/Ops leads",
    },
}

# ─────────────────────────────────────────────────────────────
# DATABASE
# ─────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS task_runs (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id    TEXT    NOT NULL,
            status     TEXT    NOT NULL,
            started_at TEXT,
            ended_at   TEXT,
            log_file   TEXT,
            is_preview INTEGER DEFAULT 0,
            note       TEXT
        );
        CREATE TABLE IF NOT EXISTS email_queue (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id     TEXT NOT NULL,
            preview_html TEXT,
            status      TEXT DEFAULT 'pending',
            created_at  TEXT,
            approved_at TEXT,
            sent_at     TEXT
        );
    """)
    conn.commit()
    conn.close()

init_db()

def get_task_status(task_id):
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM task_runs WHERE task_id=? ORDER BY id DESC LIMIT 1",
        (task_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None

def upsert_task_status(task_id, status, log_file=None, note=None, is_preview=False):
    conn = get_db()
    now = datetime.now().isoformat()
    running = conn.execute(
        "SELECT id FROM task_runs WHERE task_id=? AND status='running'",
        (task_id,)
    ).fetchone()
    if running and status in ("done", "failed"):
        conn.execute(
            "UPDATE task_runs SET status=?, ended_at=?, note=? WHERE id=?",
            (status, now, note, running["id"])
        )
    else:
        conn.execute(
            "INSERT INTO task_runs (task_id,status,started_at,log_file,is_preview) VALUES (?,?,?,?,?)",
            (task_id, status, now, log_file, int(is_preview))
        )
    conn.commit()
    conn.close()

def get_email_queue():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM email_queue WHERE status IN ('pending','approved') ORDER BY id DESC"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def queue_email(task_id, preview_html=""):
    conn = get_db()
    conn.execute(
        "INSERT INTO email_queue (task_id,preview_html,created_at) VALUES (?,?,?)",
        (task_id, preview_html, datetime.now().isoformat())
    )
    conn.commit()
    conn.close()

def update_email_status(email_id, status):
    conn = get_db()
    now = datetime.now().isoformat()
    mapping = {"approved": "approved_at", "sent": "sent_at"}
    if status in mapping:
        conn.execute(f"UPDATE email_queue SET status=?, {mapping[status]}=? WHERE id=?",
                     (status, now, email_id))
    else:
        conn.execute("UPDATE email_queue SET status=? WHERE id=?", (status, email_id))
    conn.commit()
    conn.close()

# ─────────────────────────────────────────────────────────────
# NOTEBOOK PATCHING UTILS
# ─────────────────────────────────────────────────────────────
def patch_notebook_for_preview(nb_path: str) -> str:
    """
    Clone notebook to temp file, flip SEND_EMAIL = True → False.
    Returns path to patched temp file.
    """
    with open(nb_path, encoding="utf-8") as f:
        nb = json.load(f)

    for cell in nb.get("cells", []):
        src = "".join(cell.get("source", []))
        if "SEND_EMAIL" in src:
            # Regex: SEND_EMAIL (spaces) = (spaces) True  → False
            new_src = re.sub(
                r"(SEND_EMAIL\s*=\s*)True",
                r"\g<1>False",
                src
            )
            cell["source"] = new_src
            break  # only patch first occurrence (config cell)

    tmp = tempfile.mktemp(suffix=".ipynb")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(nb, f)
    return tmp

def extract_html_from_executed_notebook(executed_nb_path: str) -> str:
    """Extract HTML display outputs from an executed notebook."""
    try:
        with open(executed_nb_path, encoding="utf-8") as f:
            nb = json.load(f)
    except Exception:
        return ""

    html_parts = []
    for cell in nb.get("cells", []):
        for output in cell.get("outputs", []):
            if output.get("output_type") in ("display_data", "execute_result"):
                html = output.get("data", {}).get("text/html", "")
                if html:
                    chunk = "".join(html) if isinstance(html, list) else html
                    # Only grab substantial HTML blocks (email-like content)
                    if len(chunk) > 200 and "<table" in chunk.lower():
                        html_parts.append(chunk)

    if not html_parts:
        return ""
    return "<hr>".join(html_parts)

# ─────────────────────────────────────────────────────────────
# TASK RUNNER
# ─────────────────────────────────────────────────────────────
def run_task(task_id: str, preview: bool = False):
    """Spawn background thread to execute task."""
    task = TASKS[task_id]
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    mode = "preview" if preview else "run"
    log_file = str(LOG_DIR / f"{task_id}_{mode}_{ts}.log")
    upsert_task_status(task_id, "running", log_file=log_file, is_preview=preview)

    def _worker():
        try:
            if task["type"] == "py":
                cmd = ["python", task["script"]]
                env = {**os.environ, **({"PREVIEW_ONLY": "1"} if preview else {})}
                with open(log_file, "w", encoding="utf-8") as lf:
                    proc = subprocess.Popen(
                        cmd, stdout=lf, stderr=subprocess.STDOUT,
                        env=env, cwd=str(CODE_DIR),
                        creationflags=0x08000000 if os.name == "nt" else 0,
                    )
                    proc.wait()
                success = proc.returncode == 0

            else:  # notebook
                nb_path = task["notebook"]
                if preview and task.get("send_var"):
                    nb_path = patch_notebook_for_preview(nb_path)

                executed_nb = str(LOG_DIR / f"{task_id}_{ts}_executed.ipynb")
                cmd = [
                    "jupyter", "nbconvert",
                    "--to", "notebook",
                    "--execute",
                    "--ExecutePreprocessor.timeout=600",
                    "--ExecutePreprocessor.kernel_name=python3",
                    nb_path,
                    "--output", executed_nb,
                ]
                with open(log_file, "w", encoding="utf-8") as lf:
                    proc = subprocess.Popen(
                        cmd, stdout=lf, stderr=subprocess.STDOUT,
                        cwd=str(CODE_DIR),
                        creationflags=0x08000000 if os.name == "nt" else 0,
                    )
                    proc.wait()
                success = proc.returncode == 0

                # Clean up temp patched notebook
                if preview and nb_path != task["notebook"]:
                    try:
                        os.remove(nb_path)
                    except Exception:
                        pass

                # Extract email HTML for preview
                if preview and task.get("layer") == 3 and success:
                    html = extract_html_from_executed_notebook(executed_nb)
                    if html:
                        preview_file = PREVIEW_DIR / f"{task_id}_preview.html"
                        preview_file.write_text(html, encoding="utf-8")
                        queue_email(task_id, html)

            upsert_task_status(task_id, "done" if success else "failed",
                               note=f"Exit {proc.returncode}" if not success else None)

        except Exception as exc:
            upsert_task_status(task_id, "failed", note=str(exc))
            with open(log_file, "a", encoding="utf-8") as lf:
                lf.write(f"\n[DASHBOARD ERROR] {exc}\n")

    threading.Thread(target=_worker, daemon=True).start()

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
STATUS_CFG = {
    "running": ("⏳", "#F57F17", "#FFFDE7", "rgba(245,127,23,0.12)"),
    "done":    ("✅", "#2E7D32", "#E8F5E9", "rgba(46,125,50,0.10)"),
    "failed":  ("❌", "#C62828", "#FFEBEE", "rgba(198,40,40,0.10)"),
    "pending": ("⏸",  "#546E7A", "#ECEFF1", "rgba(84,110,122,0.08)"),
    None:      ("○",  "#9E9E9E", "#FAFAFA", "rgba(0,0,0,0.04)"),
}

def deps_ok(task_id) -> bool:
    for dep in TASKS[task_id]["deps"]:
        s = get_task_status(dep)
        if not s or s["status"] != "done":
            return False
    return True

def output_fresh(task_id) -> bool:
    p = TASKS[task_id].get("output_check")
    if not p:
        return False
    try:
        mtime = Path(p).stat().st_mtime
        age_h = (time.time() - mtime) / 3600
        return age_h < 24
    except Exception:
        return False

def fmt_dt(iso):
    if not iso:
        return "—"
    try:
        return datetime.fromisoformat(iso).strftime("%d/%m %H:%M")
    except Exception:
        return iso

def get_latest_log(task_id):
    logs = sorted(LOG_DIR.glob(f"{task_id}_*.log"), reverse=True)
    if not logs:
        return None
    try:
        return logs[0].read_text(encoding="utf-8", errors="replace")[-6000:]
    except Exception:
        return None

def layer_tasks(layer_num):
    return [(tid, t) for tid, t in TASKS.items() if t["layer"] == layer_num]

# ─────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Global */
[data-testid="stAppViewContainer"] { background: #F8F9FA; }
.block-container { padding-top: 1rem; }

/* Layer header */
.layer-hdr {
    display: flex; align-items: center; gap: 10px;
    color: white; font-weight: 700; font-size: 13px;
    padding: 7px 14px; border-radius: 8px;
    margin: 18px 0 10px 0;
}

/* Task card */
.task-card {
    background: white;
    border: 1px solid #E0E0E0;
    border-radius: 10px;
    padding: 12px 14px 10px;
    margin-bottom: 10px;
    transition: box-shadow 0.2s;
}
.task-card:hover { box-shadow: 0 3px 12px rgba(0,0,0,0.10); }

/* Status pill */
.spill {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
    margin-left: 6px;
}

/* Metric strip */
.metrics-strip {
    display: flex; gap: 16px; flex-wrap: wrap;
    background: white; border-radius: 10px;
    padding: 12px 18px; margin-bottom: 16px;
    border: 1px solid #E0E0E0;
}
.metric-item { text-align: center; min-width: 80px; }
.metric-val  { font-size: 26px; font-weight: 800; line-height: 1.1; }
.metric-lbl  { font-size: 11px; color: #757575; }

/* Log viewer */
.log-view {
    background: #1E1E1E; color: #D4D4D4;
    font-family: "Courier New", monospace;
    font-size: 12px; line-height: 1.5;
    padding: 14px; border-radius: 8px;
    max-height: 380px; overflow-y: auto;
    white-space: pre-wrap; word-break: break-all;
}

/* Badge for email queue */
.eq-badge {
    background: #B71C1C; color: white;
    font-size: 11px; font-weight: 700;
    padding: 2px 8px; border-radius: 12px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🚀 WFM Pipeline")
    st.caption("Expedia HCM — Concentrix VN")
    st.divider()

    # Overall stats
    statuses = {tid: get_task_status(tid) for tid in TASKS}
    n_done    = sum(1 for s in statuses.values() if s and s["status"] == "done")
    n_running = sum(1 for s in statuses.values() if s and s["status"] == "running")
    n_failed  = sum(1 for s in statuses.values() if s and s["status"] == "failed")
    pending_emails = len(get_email_queue())

    c1, c2 = st.columns(2)
    c1.metric("✅ Done",    n_done)
    c2.metric("⏳ Running", n_running)
    c1.metric("❌ Failed",  n_failed)
    c2.metric("✉️ Queue",   pending_emails)

    st.divider()
    st.markdown("**⚡ Quick Actions**")

    if st.button("▶️ Run all ETL", use_container_width=True):
        for tid in ["hc_master", "iex", "rta", "atd", "performance"]:
            s = get_task_status(tid)
            if not s or s["status"] not in ("running", "done"):
                if deps_ok(tid):
                    run_task(tid)
        st.toast("ETL pipeline started")

    if st.button("🤖 Run all Bots", use_container_width=True):
        for tid in ["bot_atd", "bot_ic", "bot_lc", "bot_ou", "bot_ic_summary"]:
            run_task(tid)
        st.toast("All bots started")

    if st.button("👁 Preview all Emails", use_container_width=True, type="primary"):
        email_tasks = [tid for tid, t in TASKS.items() if t["layer"] == 3]
        count = 0
        for tid in email_tasks:
            if deps_ok(tid):
                run_task(tid, preview=True)
                count += 1
        st.toast(f"Generating {count} email previews...")

    st.divider()
    auto_refresh = st.toggle("🔄 Auto-refresh 10s", value=False)
    if st.button("🔄 Refresh", use_container_width=True):
        st.rerun()

# ─────────────────────────────────────────────────────────────
# MAIN TABS
# ─────────────────────────────────────────────────────────────
tab_pipeline, tab_email, tab_logs, tab_history = st.tabs([
    "🗂 Pipeline",
    f"✉️ Email Queue {'🔴' if pending_emails else ''}",
    "📋 Logs",
    "📊 History",
])

# ══════════════════════════════════════════════════════════════
# TAB 1 — PIPELINE
# ══════════════════════════════════════════════════════════════
with tab_pipeline:
    st.markdown("### 🗂 Pipeline Status")

    for layer_num in sorted(LAYERS.keys()):
        icon, layer_name, color = LAYERS[layer_num]
        tasks = layer_tasks(layer_num)

        st.markdown(
            f'<div class="layer-hdr" style="background:{color}">'
            f'{icon}&nbsp; Layer {layer_num} — {layer_name} '
            f'<span style="font-weight:400;font-size:11px;opacity:.8">'
            f'({len(tasks)} tasks)</span>'
            f'</div>',
            unsafe_allow_html=True
        )

        cols = st.columns(3)
        for idx, (tid, tdata) in enumerate(tasks):
            with cols[idx % 3]:
                s      = get_task_status(tid)
                s_key  = s["status"] if s else None
                s_icon, s_color, s_bg, _ = STATUS_CFG.get(s_key, STATUS_CFG[None])
                is_running = s_key == "running"
                can_run    = not is_running
                is_email   = tdata["layer"] == 3

                # Card
                with st.container():
                    st.markdown(
                        f'<div class="task-card" style="border-left:4px solid {color}">'
                        f'<b style="font-size:13px">{tdata["name"]}</b>'
                        f'<span class="spill" style="background:{s_bg};color:{s_color}">'
                        f'{s_icon} {s_key or "never"}</span>'
                        f'<div style="font-size:11px;color:#757575;margin:4px 0 6px">'
                        f'{tdata["desc"]}</div>',
                        unsafe_allow_html=True
                    )

                    # Dependency & output info
                    if tdata["deps"]:
                        dep_names = [TASKS[d]["name"] for d in tdata["deps"] if d in TASKS]
                        met = all(
                            get_task_status(d) and get_task_status(d)["status"] == "done"
                            for d in tdata["deps"]
                        )
                        color_dep = "#2E7D32" if met else "#C62828"
                        st.markdown(
                            f'<div style="font-size:10px;color:{color_dep};margin-bottom:4px">'
                            f'{"✔" if met else "✘"} Deps: {", ".join(dep_names)}</div>',
                            unsafe_allow_html=True
                        )

                    if s:
                        st.markdown(
                            f'<div style="font-size:10px;color:#9E9E9E">Last: {fmt_dt(s.get("started_at"))}'
                            f'{"  →  " + fmt_dt(s.get("ended_at")) if s.get("ended_at") else ""}</div>',
                            unsafe_allow_html=True
                        )

                    st.markdown('</div>', unsafe_allow_html=True)
                    st.markdown("")

                    # Buttons
                    if is_email:
                        b1, b2, b3 = st.columns(3)
                        with b1:
                            if st.button("👁 Preview", key=f"prev_{tid}",
                                         disabled=not can_run or not deps_ok(tid),
                                         use_container_width=True):
                                run_task(tid, preview=True)
                                st.toast(f"Generating preview: {tdata['name']}")
                        with b2:
                            preview_exists = (PREVIEW_DIR / f"{tid}_preview.html").exists()
                            if st.button("📋 Queue", key=f"q_{tid}",
                                         disabled=not preview_exists,
                                         use_container_width=True):
                                html = (PREVIEW_DIR / f"{tid}_preview.html").read_text(encoding="utf-8")
                                queue_email(tid, html)
                                st.toast(f"Added to queue: {tdata['name']}")
                        with b3:
                            if st.button("▶️ Send", key=f"run_{tid}",
                                         disabled=not can_run or not deps_ok(tid),
                                         use_container_width=True, type="primary"):
                                run_task(tid, preview=False)
                                st.toast(f"Sending: {tdata['name']}")
                    else:
                        b1, b2 = st.columns(2)
                        with b1:
                            if st.button(
                                "⏹ Running…" if is_running else "▶️ Run",
                                key=f"run_{tid}",
                                disabled=is_running or not deps_ok(tid),
                                use_container_width=True,
                                type="primary" if not is_running else "secondary",
                            ):
                                run_task(tid)
                                st.toast(f"Started: {tdata['name']}")
                        with b2:
                            fresh = output_fresh(tid)
                            if tdata.get("output_check"):
                                st.markdown(
                                    f'<div style="font-size:11px;text-align:center;'
                                    f'color:{"#2E7D32" if fresh else "#9E9E9E"};padding-top:6px">'
                                    f'{"📦 Fresh" if fresh else "📭 Stale"}</div>',
                                    unsafe_allow_html=True
                                )

                st.markdown("---")

# ══════════════════════════════════════════════════════════════
# TAB 2 — EMAIL QUEUE
# ══════════════════════════════════════════════════════════════
with tab_email:
    st.markdown("### ✉️ Email Approval Queue")
    st.caption("Preview email trước khi gửi. Approve → Send để gửi thật qua Outlook.")

    queue = get_email_queue()

    if not queue:
        st.info(
            "📭 Queue trống. Vào tab **Pipeline** → nhấn **👁 Preview** cho email task bất kỳ "
            "để generate preview, sau đó **📋 Queue** để đưa vào đây."
        )
    else:
        for item in queue:
            tid   = item["task_id"]
            eid   = item["id"]
            name  = TASKS.get(tid, {}).get("name", tid)
            stat  = item["status"]
            creat = fmt_dt(item.get("created_at"))

            badge = "🟡 Pending" if stat == "pending" else "🟢 Approved"
            with st.expander(f"{badge}  ·  {name}  ·  {creat}", expanded=True):
                # Load preview HTML
                html_content = item.get("preview_html", "")
                if not html_content:
                    pf = PREVIEW_DIR / f"{tid}_preview.html"
                    if pf.exists():
                        html_content = pf.read_text(encoding="utf-8")

                if html_content:
                    st.markdown("**📧 Email Preview:**")
                    st.components.v1.html(html_content, height=520, scrolling=True)
                else:
                    st.warning("⚠️ Preview HTML chưa có — nhấn **👁 Preview** lại từ tab Pipeline.")

                st.markdown("---")
                ca, cb, cc, _ = st.columns([1, 1, 2, 1])
                with ca:
                    if stat == "pending":
                        if st.button("✅ Approve", key=f"app_{eid}",
                                     type="primary", use_container_width=True):
                            update_email_status(eid, "approved")
                            st.success("Approved!")
                            st.rerun()
                with cb:
                    if st.button("❌ Reject", key=f"rej_{eid}", use_container_width=True):
                        update_email_status(eid, "rejected")
                        st.rerun()
                with cc:
                    if stat == "approved":
                        if st.button(f"📨 Send — {name}", key=f"send_{eid}",
                                     type="primary", use_container_width=True):
                            run_task(tid, preview=False)
                            update_email_status(eid, "sent")
                            st.success(f"✅ Sent: {name}")
                            time.sleep(1)
                            st.rerun()
                    else:
                        st.info("Approve trước rồi mới có thể Send", icon="ℹ️")

    # Sent history
    st.divider()
    with st.expander("📜 Sent History (last 20)"):
        conn = get_db()
        sent = conn.execute(
            "SELECT * FROM email_queue WHERE status IN ('sent','rejected') ORDER BY id DESC LIMIT 20"
        ).fetchall()
        conn.close()
        if sent:
            import pandas as pd
            df = pd.DataFrame([dict(r) for r in sent])
            df["task_id"] = df["task_id"].map(lambda x: TASKS.get(x, {}).get("name", x))
            st.dataframe(
                df[["task_id","status","created_at","approved_at","sent_at"]],
                use_container_width=True
            )
        else:
            st.caption("Chưa có email nào được gửi.")

# ══════════════════════════════════════════════════════════════
# TAB 3 — LOGS
# ══════════════════════════════════════════════════════════════
with tab_logs:
    st.markdown("### 📋 Live Logs")

    selected = st.selectbox(
        "Chọn task",
        options=list(TASKS.keys()),
        format_func=lambda x: f"[L{TASKS[x]['layer']}] {TASKS[x]['name']}",
    )

    s = get_task_status(selected)
    if s:
        c1, c2, c3, c4 = st.columns(4)
        icon = STATUS_CFG.get(s["status"], STATUS_CFG[None])[0]
        c1.metric("Status",  f"{icon} {s['status']}")
        c2.metric("Started", fmt_dt(s.get("started_at")))
        c3.metric("Ended",   fmt_dt(s.get("ended_at")))
        c4.metric("Preview", "Yes" if s.get("is_preview") else "No")

    log_content = get_latest_log(selected)
    if log_content:
        # Colorize errors in log
        display_log = log_content
        st.markdown(
            f'<div class="log-view">{display_log}</div>',
            unsafe_allow_html=True
        )
        if st.button("📋 Copy log path"):
            logs = sorted(LOG_DIR.glob(f"{selected}_*.log"), reverse=True)
            if logs:
                st.code(str(logs[0]))
    else:
        st.info("Chưa có log. Chạy task để xem output ở đây.")

    if s and s.get("status") == "running":
        st.warning("⏳ Task đang chạy — nhấn Refresh để cập nhật log")
        if st.button("🔄 Refresh log"):
            st.rerun()

# ══════════════════════════════════════════════════════════════
# TAB 4 — HISTORY
# ══════════════════════════════════════════════════════════════
with tab_history:
    st.markdown("### 📊 Run History")

    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM task_runs ORDER BY id DESC LIMIT 100"
    ).fetchall()
    conn.close()

    if rows:
        import pandas as pd
        df = pd.DataFrame([dict(r) for r in rows])
        df["task_name"] = df["task_id"].map(lambda x: TASKS.get(x, {}).get("name", x))
        df["layer"] = df["task_id"].map(lambda x: f"L{TASKS.get(x, {}).get('layer', '?')}")
        df["mode"] = df["is_preview"].map({1: "preview", 0: "run"})

        # Filter controls
        cf1, cf2, cf3 = st.columns(3)
        with cf1:
            status_filter = st.multiselect(
                "Status", ["done","failed","running","pending"],
                default=["done","failed","running"]
            )
        with cf2:
            layer_filter = st.multiselect(
                "Layer", [0,1,2,3], default=[0,1,2,3],
                format_func=lambda x: f"L{x} {LAYERS[x][1]}"
            )
        with cf3:
            mode_filter = st.multiselect("Mode", ["run","preview"], default=["run","preview"])

        df_filt = df[
            df["status"].isin(status_filter) &
            df["task_id"].map(lambda x: TASKS.get(x, {}).get("layer", -1)).isin(layer_filter) &
            df["mode"].isin(mode_filter)
        ]

        st.dataframe(
            df_filt[["id","layer","task_name","mode","status","started_at","ended_at","note"]],
            use_container_width=True,
            height=400,
        )

        # Summary chart
        st.divider()
        st.markdown("**Success rate by task (last 100 runs)**")
        summary = (
            df.groupby(["task_name","status"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        st.bar_chart(
            summary.set_index("task_name")[["done","failed"]] if "failed" in summary.columns
            else summary.set_index("task_name"),
            color=["#2E7D32","#C62828"],
        )
    else:
        st.info("Chưa có run history nào. Chạy task đầu tiên đi!")

# ─────────────────────────────────────────────────────────────
# AUTO-REFRESH
# ─────────────────────────────────────────────────────────────
if auto_refresh:
    time.sleep(10)
    st.rerun()
