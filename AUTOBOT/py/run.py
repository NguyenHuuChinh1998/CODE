# %%
import subprocess
import sys
import time
from datetime import datetime

PY_DIR = r"C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\BI_Task\CODE\Python_Code\AUTOBOT\py"

RUN_ORDER = [
    "download.py",
    "atd_realtime.py",
    "ic_monitor.py",
    "lc_flag.py",
    "ou_present.py",
]

import os
bot_start = datetime.now()
print(f"\n{'═'*55}")
print(f"🚀 AUTOBOT started: {bot_start.strftime('%d-%b-%Y %H:%M:%S')}")
print(f"{'═'*55}")

results = []

for script in RUN_ORDER:
    path   = os.path.join(PY_DIR, script)
    label  = script.replace(".py", "").upper()
    t0     = time.time()
    ts     = datetime.now().strftime("%H:%M:%S")

    print(f"\n▶  [{ts}] Running: {script}")
    print(f"{'─'*55}")

    try:
        proc = subprocess.run(
            [sys.executable, path],
            capture_output=False,   # print stdout/stderr trực tiếp ra console
            text=True,
        )
        elapsed = round(time.time() - t0, 1)
        status  = "✅ OK" if proc.returncode == 0 else f"❌ Exit {proc.returncode}"
    except Exception as e:
        elapsed = round(time.time() - t0, 1)
        status  = f"❌ ERROR: {e}"

    results.append((label, elapsed, status))
    print(f"{'─'*55}")
    print(f"   {status} | {label} finished in {elapsed}s")

# ── Summary ───────────────────────────────────────────────────
total = round((datetime.now() - bot_start).total_seconds(), 1)
print(f"\n{'═'*55}")
print(f"📋 SUMMARY")
print(f"{'─'*55}")
for label, elapsed, status in results:
    print(f"  {status:<18} {label:<20} {elapsed:>6}s")
print(f"{'─'*55}")
print(f"  ⏱  Total time : {total}s  ({round(total/60,1)} min)")
print(f"  🏁 Finished   : {datetime.now().strftime('%d-%b-%Y %H:%M:%S')}")
print(f"{'═'*55}\n")


