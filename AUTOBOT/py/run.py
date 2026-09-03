# %%
import subprocess
import sys
import time
import os
import requests
import json
from datetime import datetime, timedelta

CREATE_NO_WINDOW = 0x08000000

PROFILE_PATH    = r"C:/temp/new_chrome_profile"
PROFILE_KEYWORD = "new_chrome_profile"
PY_DIR          = r"C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\BI_Task\CODE\Python_Code\AUTOBOT\py"
TEAMS_WEBHOOK_URL = "https://default599e51d62f8c43478e591f795a51a9.8c.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/c24f30c010df45a6a6dac9421643bb34/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=5vWDl18a7-IWSvHuZAWgGtQcwM54nEapSArj4JVPnGg"
BASE_CAPTURE    = r"C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\Rawdata\CAPTURE"
LAST_ALERT_FILE = r"C:\temp\last_download_alert.txt"

RUN_ORDER = [
    "download_test.py",
    "atd_realtime.py",
    "ic_monitor.py",
    "lc_flag.py",
    "ou_present.py",
    "ic_summary.py"
]
DOWNLOAD_SCRIPT = "download_test.py"
PS_TIMEOUT      = 15


def _run_ps(cmd, timeout=PS_TIMEOUT):
    try:
        return subprocess.run(
            ["powershell", "-NoProfile", "-Command", cmd],
            capture_output=True, text=True,
            timeout=timeout
        )
    except subprocess.TimeoutExpired:
        print(f"  ⚠️ PowerShell command timed out after {timeout}s — skipping")
        return None
    except Exception as e:
        print(f"  ⚠️ PowerShell error: {e}")
        return None


def kill_self_previous():
    """Kill any previously hung running_bot process to allow clean restart."""
    current_pid = os.getpid()
    cmd = (
        f"Get-CimInstance Win32_Process -Filter \"name='python.exe'\" | "
        f"Where-Object {{ $_.CommandLine -like '*running_bot*' "
        f"-and $_.ProcessId -ne {current_pid} }} | "
        f"ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force }}"
    )
    r = _run_ps(cmd)
    if r and r.returncode == 0:
        print("  ✅ Killed previous hung running_bot instance")
    else:
        print("  ℹ️  No previous running_bot found")


def kill_orphan_chromedriver(profile_keyword=PROFILE_KEYWORD):
    try:
        r1 = subprocess.run(
            ["taskkill", "/F", "/IM", "chromedriver.exe"],
            capture_output=True, text=True,
            timeout=10
        )
        print("✅ Killed chromedriver" if "SUCCESS" in r1.stdout
              else "ℹ️  No chromedriver found")
    except subprocess.TimeoutExpired:
        print("⚠️ taskkill chromedriver timed out")
    except Exception as e:
        print(f"⚠️ chromedriver: {e}")

    cmd = (
        f"Get-CimInstance Win32_Process -Filter \"name='chrome.exe'\" | "
        f"Where-Object {{ $_.CommandLine -like '*{profile_keyword}*' }} | "
        f"ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force }}"
    )
    r2 = _run_ps(cmd)
    if r2:
        print(f"✅ Killed bot Chrome ({profile_keyword})" if r2.returncode == 0
              else "ℹ️  No bot Chrome found")

    lock_files = [
        os.path.join(PROFILE_PATH, "SingletonLock"),
        os.path.join(PROFILE_PATH, "SingletonSocket"),
        os.path.join(PROFILE_PATH, "SingletonCookie"),
        os.path.join(PROFILE_PATH, "Default", "LOCK"),
        os.path.join(PROFILE_PATH, "Default", "lockfile"),
    ]
    for lf in lock_files:
        try:
            if os.path.exists(lf):
                os.remove(lf)
                print(f"  ✅ Removed lock: {os.path.basename(lf)}")
        except Exception as e:
            print(f"  ⚠️ Could not remove {os.path.basename(lf)}: {e}")

    time.sleep(2)


def kill_orphan_bots(scripts=None):
    if scripts is None:
        scripts = RUN_ORDER
    for script in scripts:
        cmd = (
            f"Get-CimInstance Win32_Process -Filter \"name='python.exe'\" | "
            f"Where-Object {{ $_.CommandLine -like '*{script}*' }} | "
            f"ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force }}"
        )
        r = _run_ps(cmd)
        if r is None:
            print(f"  ⚠️ Timeout checking: {script}")
        elif "successfully" in (r.stdout or "").lower():
            print(f"  ✅ Killed hanging: {script}")
        else:
            print(f"  ℹ️  No hanging process: {script}")


def send_warning_webhook(msg, cooldown_minutes=25):
    now = datetime.now()
    if os.path.exists(LAST_ALERT_FILE):
        try:
            last = datetime.fromisoformat(open(LAST_ALERT_FILE).read().strip())
            if (now - last).total_seconds() < cooldown_minutes * 60:
                print(f"⏭️  Alert suppressed (cooldown {cooldown_minutes}min)")
                return
        except:
            pass

    payload = {"html": (
        f'<p><b style="color:#c0392b;font-size:16px;">⚠️ BOT DOWNLOAD ALERT</b><br>'
        f'<span style="font-size:13px;">{msg}<br>'
        f'Time: <b>{now.strftime("%d-%b-%Y %H:%M")}</b></span></p>'
    )}
    try:
        r = requests.post(TEAMS_WEBHOOK_URL,
                          headers={"Content-Type": "application/json"},
                          data=json.dumps(payload), timeout=30)
        print(f"⚠️ Warning sent: {r.status_code}")
        open(LAST_ALERT_FILE, 'w').write(now.isoformat())
    except Exception as e:
        print(f"⚠️ Webhook failed: {e}")


def check_download_success(start_time):
    CHECK_DIRS = {
        "current_interval": os.path.join(BASE_CAPTURE, "current_interval"),
        "current_agent"   : os.path.join(BASE_CAPTURE, "current_agent"),
        "lc_rawdata"      : os.path.join(BASE_CAPTURE, "lc_rawdata_in_console"),
    }
    missing = []
    for name, folder in CHECK_DIRS.items():
        try:
            new_files = [
                f for f in os.listdir(folder)
                if os.path.isfile(os.path.join(folder, f))
                   and not f.startswith('_')
                   and datetime.fromtimestamp(
                       os.path.getmtime(os.path.join(folder, f))
                   ) >= start_time
            ]
            if new_files:
                print(f"  ✅ {name}: {len(new_files)} new file(s) → {new_files[-1]}")
            else:
                print(f"  ❌ {name}: no new file since {start_time.strftime('%H:%M:%S')}")
                missing.append(name)
        except Exception as e:
            print(f"  ⚠️ {name}: check error — {e}")
            missing.append(name)
    return missing


bot_start = datetime.now()
print(f"\n{'═'*55}")
print(f"🚀 AUTOBOT started: {bot_start.strftime('%d-%b-%Y %H:%M:%S')}")
print(f"{'═'*55}")

print("\n🧹 Cleaning up orphaned processes...")
kill_self_previous()
kill_orphan_bots()
kill_orphan_chromedriver()

results     = []
download_ok = True

for script in RUN_ORDER:
    path  = os.path.join(PY_DIR, script)
    label = script.replace(".py", "").upper()
    t0    = time.time()
    ts    = datetime.now().strftime("%H:%M:%S")

    if not download_ok and script != DOWNLOAD_SCRIPT:
        print(f"\n⏭️  [{ts}] Skipping {script} — download failed")
        results.append((label, 0.0, "⏭️ SKIPPED"))
        continue

    print(f"\n▶  [{ts}] Running: {script}")
    print(f"{'─'*55}")

    download_start = datetime.now()

    try:
        proc = subprocess.run(
            [sys.executable, path],
            capture_output=False, text=True,
            creationflags=CREATE_NO_WINDOW,
            timeout=600,
        )
        elapsed = round(time.time() - t0, 1)
        status  = "✅ OK" if proc.returncode == 0 else f"❌ Exit {proc.returncode}"

    except subprocess.TimeoutExpired:
        elapsed = round(time.time() - t0, 1)
        status  = "❌ TIMEOUT (>10min)"
        print(f"  ❌ {script} timed out — killing all bot processes")
        kill_orphan_bots()
        kill_orphan_chromedriver()

    except Exception as e:
        elapsed = round(time.time() - t0, 1)
        status  = f"❌ ERROR: {e}"

    if script == DOWNLOAD_SCRIPT:
        print(f"\n🔍 Verifying files moved after {download_start.strftime('%H:%M:%S')}...")
        missing = check_download_success(start_time=download_start)
        if missing:
            download_ok = False
            msg = (
                f"Data could not be downloaded at this time, please check !<br>"
                f"Folders missing new files: <b>{', '.join(missing)}</b>"
            )
            send_warning_webhook(msg)
            status = f"⚠️ Download incomplete — missing: {', '.join(missing)}"
        else:
            print(f"  ✅ All folders verified — files moved successfully")

    results.append((label, elapsed, status))
    print(f"{'─'*55}")
    print(f"   {status} | {label} finished in {elapsed}s")

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


