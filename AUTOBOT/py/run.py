# %%
import subprocess
import sys
import time
import os
from datetime import datetime

#subprocess.run(["shutdown", "/s", "/t", "0"])
CREATE_NO_WINDOW = 0x08000000  # Windows flag ẩn console

PROFILE_PATH  = r"C:/temp/new_chrome_profile"
PROFILE_KEYWORD = "new_chrome_profile"
PY_DIR = r"C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\BI_Task\CODE\Python_Code\AUTOBOT\py"

RUN_ORDER = [
    "download.py",
    "atd_realtime.py",
    "ic_monitor.py",
    "lc_flag.py",
    "ou_present.py",
]

def kill_orphan_chromedriver(profile_keyword=PROFILE_KEYWORD):
    # Kill chromedriver
    try:
        r1 = subprocess.run(["taskkill", "/F", "/IM", "chromedriver.exe"],
                            capture_output=True, text=True)
        print("✅ Killed chromedriver" if "SUCCESS" in r1.stdout
              else "ℹ️  No chromedriver found")
    except Exception as e:
        print(f"⚠️ chromedriver: {e}")

    # Kill bot Chrome
    try:
        ps_cmd = (
            f"Get-WmiObject Win32_Process -Filter \"name='chrome.exe'\" | "
            f"Where-Object {{ $_.CommandLine -like '*{profile_keyword}*' }} | "
            f"ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force }}"
        )
        r2 = subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd],
                            capture_output=True, text=True)
        print(f"✅ Killed bot Chrome ({profile_keyword})" if r2.returncode == 0
              else "ℹ️  No bot Chrome found")
    except Exception as e:
        print(f"⚠️ Chrome: {e}")

    # Remove Chrome profile lock files
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

    time.sleep(3)

bot_start = datetime.now()
print(f"\n{'═'*55}")
print(f"🚀 AUTOBOT started: {bot_start.strftime('%d-%b-%Y %H:%M:%S')}")
print(f"{'═'*55}")

print("\n🧹 Cleaning up orphaned processes...")
kill_orphan_chromedriver()

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
            capture_output=False,
            text=True,
            creationflags=CREATE_NO_WINDOW,
        )
        elapsed = round(time.time() - t0, 1)
        status  = "✅ OK" if proc.returncode == 0 else f"❌ Exit {proc.returncode}"
    except Exception as e:
        elapsed = round(time.time() - t0, 1)
        status  = f"❌ ERROR: {e}"

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


