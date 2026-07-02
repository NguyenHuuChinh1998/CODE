# %%
import glob
import io
import os
import shutil
import sys
import time
import pathlib
import re
from collections import OrderedDict
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

import numpy as np
import pandas as pd
import polars as pl
import pyautogui
import win32clipboard

from PIL import Image, ImageGrab

from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import (
    NoAlertPresentException,
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException
)

# %%
def input_ou(folder_path: str, sheet_name: str | None = None):
    import os, re, glob
    from datetime import datetime
    import polars as pl

    DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    DAY2OFFSET = {d: i for i, d in enumerate(DAYS)}

    files = glob.glob(os.path.join(folder_path, "*.xlsx")) + glob.glob(os.path.join(folder_path, "*.csv"))
    if not files:
        return pl.DataFrame(
            schema={
                "PST_Date": pl.Date, "VN_Date": pl.Date,
                "PST_Datetime": pl.Datetime, "VN_Datetime": pl.Datetime,
                "PST_Intervals": pl.Utf8, "VN_Intervals": pl.Utf8,
                "ou_field": pl.Utf8, "value": pl.Float64
            }
        )
    out_frames = []
    for f in sorted(files):
        m = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(f))
        if not m:
            continue
        week_start = datetime.strptime(m.group(1), "%Y-%m-%d").date()
        if f.lower().endswith(".xlsx"):
            df = pl.read_excel(f, sheet_name=sheet_name)
        else:
            try:
                df = pl.read_csv(f, encoding="utf-8")
            except Exception:
                df = pl.read_csv(f, encoding="ISO-8859-1", ignore_errors=True)

        if df.width == 0 or df.height == 0:
            continue
        cols = df.columns
        head = df.head(1)
        top = {c: ("" if (v := head.select(pl.col(c)).to_series()[0]) is None else str(v)).strip().strip('"').strip()
               for c in cols}
        rename_map = {}
        if len(cols) >= 1: rename_map[cols[0]] = "PST_Intervals"
        if len(cols) >= 2: rename_map[cols[1]] = "VN_Intervals"
        diff_cols = []
        for i, c in enumerate(cols):
            if c in DAYS:
                lab = top.get(c, "").upper()
                if lab == "REQ W":
                    rename_map[c] = f"{c} Req W"
                elif lab == "PROV":
                    rename_map[c] = f"{c} Prov"
                elif lab == "DIFF":
                    diff_cols.append(c)

                for g in cols[i+1:i+3]:
                    lab2 = top.get(g, "").upper()
                    if lab2 == "REQ W":
                        rename_map[g] = f"{c} Req W"
                    elif lab2 == "PROV":
                        rename_map[g] = f"{c} Prov"
                    elif lab2 == "DIFF":
                        diff_cols.append(g)

        df = df.rename(rename_map)

        df = df.slice(1)

        drop_cols = [c for c in ("VN_Intervals",) if c in df.columns] + [c for c in diff_cols if c in df.columns]
        if drop_cols:
            df = df.drop(drop_cols, strict=False)

        if "PST_Intervals" in df.columns:
            df = df.with_columns([
                pl.col("PST_Intervals").cast(pl.Float64, strict=False).alias("_pst_num"),
                pl.col("PST_Intervals").cast(pl.Utf8,   strict=False).alias("_pst_str"),
            ])
            minutes_num = (pl.col("_pst_num") * 1440).round(0).cast(pl.Int64)
            hh_num = (minutes_num // 60) % 24
            mm_num = minutes_num % 60
            hh_str = pl.col("_pst_str").str.extract(r"(\d{1,2}):\d{2}", 1).cast(pl.Int64, strict=False)
            mm_str = pl.col("_pst_str").str.extract(r"\d{1,2}:(\d{2})", 1).cast(pl.Int64, strict=False)
            df = df.with_columns([
                pl.coalesce([hh_num, hh_str]).alias("_hh"),
                pl.coalesce([mm_num, mm_str]).alias("_mm"),
            ])
            df = df.with_columns(
                (pl.col("_hh").cast(pl.Utf8).str.pad_start(2, "0") + pl.lit(":") +
                 pl.col("_mm").cast(pl.Utf8).str.pad_start(2, "0")).alias("PST_Intervals")
            ).drop(["_pst_num", "_pst_str"])
        day_cols = [c for d in DAYS for c in (f"{d} Req W", f"{d} Prov") if c in df.columns]
        if not day_cols:
            continue
        df = df.with_columns([pl.col(c).str.replace_all('"', "").cast(pl.Float64, strict=False).alias(c) for c in day_cols])
        long_df = df.melt(id_vars=["PST_Intervals"], value_vars=day_cols, variable_name="day_field", value_name="value")
        long_df = long_df.with_columns([
            pl.col("day_field").str.replace(r"\s+(Req W|Prov)$", "").alias("day"),
            pl.col("day_field").str.extract(r"(Req W|Prov)$", 1).alias("ou_field")
        ]).drop("day_field")
        long_df = long_df.with_columns(pl.col("day").replace(DAY2OFFSET).cast(pl.Int32).alias("offset"))
        long_df = long_df.with_columns(
            (pl.lit(week_start).cast(pl.Date) + pl.duration(days=pl.col("offset").cast(pl.Int64))).alias("PST_Date")
        ).drop("offset")
        long_df = long_df.filter(pl.col("ou_field") == "Req W").with_columns(pl.col("value").round(2).alias("value"))
        if "_hh" not in long_df.columns or "_mm" not in long_df.columns:
            long_df = long_df.with_columns([
                pl.col("PST_Intervals").str.slice(0, 2).cast(pl.Int64, strict=False).alias("_hh"),
                pl.col("PST_Intervals").str.slice(3, 2).cast(pl.Int64, strict=False).alias("_mm"),
            ])
        long_df = long_df.with_columns(pl.col("PST_Date").cast(pl.Datetime).alias("_base_dt"))
        long_df = long_df.with_columns(
            (pl.col("_base_dt") + pl.duration(hours=pl.col("_hh"), minutes=pl.col("_mm"))).alias("PST_Datetime")
        ).drop(["_base_dt", "_hh", "_mm"])

        long_df = long_df.with_columns(
            (pl.col("PST_Datetime") + pl.duration(hours=14)).alias("VN_Datetime")
        )
        long_df = long_df.with_columns([
            pl.col("VN_Datetime").dt.date().alias("VN_Date"),
            pl.col("VN_Datetime").dt.strftime("%H:%M").alias("VN_Intervals"),
        ])
        long_df = long_df.select([
            "PST_Date", "VN_Date", "PST_Datetime", "VN_Datetime",
            "PST_Intervals", "VN_Intervals", "ou_field", "value"
        ])

        out_frames.append(long_df)

    if not out_frames:
        return pl.DataFrame(
            schema={
                "PST_Date": pl.Date, "VN_Date": pl.Date,
                "PST_Datetime": pl.Datetime, "VN_Datetime": pl.Datetime,
                "PST_Intervals": pl.Utf8, "VN_Intervals": pl.Utf8, "ou_field": pl.Utf8, "value": pl.Float64
            }
        )

    return pl.concat(out_frames, how="vertical")

# %%
def input_data(data_dir: str) -> pl.DataFrame:
    per_file_agg: list[pl.DataFrame] = []

    for path in pathlib.Path(data_dir).glob("**/*.xlsx"):
        try:
            df = pl.read_excel(path)
        except Exception:
            continue
        if df.is_empty():
            continue
        need = ["Intervals", "Work Category", "Duration", "Scheduled Activity"]
        name_map = {}
        for want in need:
            if want not in df.columns:
                for c in df.columns:
                    if c.strip().lower() == want.lower():
                        name_map[c] = want
                        break
        if name_map:
            df = df.rename(name_map)
        if "Intervals" not in df.columns:
            df = df.with_columns(pl.lit(None).cast(pl.Datetime).alias("Intervals"))
        if "Work Category" not in df.columns:
            df = df.with_columns(pl.lit(None).cast(pl.Utf8).alias("Work Category"))
        if "Scheduled Activity" not in df.columns:
            df = df.with_columns(pl.lit(None).cast(pl.Utf8).alias("Scheduled Activity"))
        if "Duration" not in df.columns:
            df = df.with_columns(pl.lit(0.0).alias("Duration"))

        df = df.with_columns([
            pl.col("Intervals").cast(pl.Datetime, strict=False).alias("VN_Datetime"),
            pl.col("Work Category").cast(pl.Utf8).str.to_lowercase()
                .str.replace_all(r"^\s+|\s+$", "").alias("_wc"),
            pl.col("Scheduled Activity").cast(pl.Utf8).str.replace_all(r"^\s+|\s+$","").alias("_sa"),
            pl.col("Duration").cast(pl.Float64, strict=False).fill_null(0).alias("_dur")  # giây
        ])

        per = (
            df.group_by("VN_Datetime")
              .agg([
                  (pl.when(pl.col("_wc") == "productive")
                     .then(pl.col("_dur")).otherwise(0)).sum().alias("_hc_sec"),
                  (pl.when(pl.col("_sa").is_in(["No Call/No Show", "Termination"]))
                     .then(pl.col("_dur")).otherwise(0)).sum().alias("_upl_sec"),
                  (pl.when(pl.col("_sa").is_in(["PTO", "Paid Leave"]))
                     .then(pl.col("_dur")).otherwise(0)).sum().alias("_pl_sec"),
                  (pl.when(pl.col("_sa").is_in(["Lunch", "Break_1", "Break_2"]))
                     .then(pl.col("_dur")).otherwise(0)).sum().alias("_meal_break_sec"),
              ])
              .with_columns([
                  (pl.col("_hc_sec")/1800).alias("Scheduled"),
                  (pl.col("_upl_sec")/1800).alias("UPL"),
                  (pl.col("_pl_sec")/1800).alias("PL"),
                  (pl.col("_meal_break_sec")/1800).alias("Meal|Break"),
              ])
              .select(["VN_Datetime", "Scheduled", "UPL", "PL", "Meal|Break"])
        )

        if per.height:
            per_file_agg.append(per)

    if not per_file_agg:
        return pl.DataFrame(schema={
            "VN_Datetime": pl.Datetime, "Scheduled": pl.Float64,
            "UPL": pl.Float64, "PL": pl.Float64
        })
    return (
        pl.concat(per_file_agg, how="vertical")
          .group_by("VN_Datetime")
          .agg([
              pl.col("Scheduled").sum().cast(pl.Float64, strict=False).round(2).alias("Scheduled"),
              pl.col("UPL").sum().alias("UPL"),
              pl.col("PL").sum().alias("PL"),
              pl.col("Meal|Break").sum().cast(pl.Float64, strict=False).round(2).alias("Meal|Break"),
          ])
          .sort("VN_Datetime")
    )

# %%
first_glob_1 = "C:/Users/huuchinh.nguyen"
first_glob_2 = "C:/Users/ADMIN"

if os.path.exists(first_glob_1):
    first_glob = first_glob_1
elif os.path.exists(first_glob_2):
    first_glob = first_glob_2
else:
    raise FileNotFoundError(f"Neither {first_glob_1} nor {first_glob_2} exists.")

folder_paths = {
    "ou":f'{first_glob}/Concentrix Corporation/WFM-Expedia-HCM - Branding files/Rawdata/INPUT_OU_MAIL',
    "output_iex_intervals":f'{first_glob}/Concentrix Corporation/WFM-Expedia-HCM - Branding files/Rawdata/OUTPUT_AGENT_IEX_INTERVALS',
}

OU_Raw = input_ou(folder_paths["ou"])
Schedule_Raw = input_data(folder_paths["output_iex_intervals"])
OU_Raw

# %%
expedia_url = "https://console.vap.expedia.com/analytics-console-user-interface/optics/agentRealtime"
teams_url = "https://teams.microsoft.com/"
# group_chat_name = "Huu Chinh Nguyen"
parsed_url = urlparse(expedia_url)
path_fragment = parsed_url.path.split('/')[-1]
report_name = "Realtime Staffing Monitoring"
wait_seconds = 10
start_time = datetime.now()

chrome_options = Options()
chrome_options.add_argument(r'--user-data-dir=C:/temp/new_chrome_profile')
chrome_options.add_argument(r'--profile-directory=Default')
chrome_options.add_argument("--start-maximized")

service = Service(ChromeDriverManager().install())


service = Service(r"C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\Rawdata\CAPTURE\chromedriver-win64\chromedriver.exe") 
driver = webdriver.Chrome(service=service, options=chrome_options)


def send_to_clipboard(image):
    output = io.BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()
    print("Image copied to clipboard!")
    

wait = WebDriverWait(driver, 15)
driver.get(expedia_url)

def check_and_login(driver, expedia_url, wait_time=10):
    driver.get(expedia_url)
    time.sleep(10)

    try:
        sign_in_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="console-okta-sign-in"]'))
        )
        print("🔑 Sign-in required detected! Clicking...")
        sign_in_button.click()
        time.sleep(5)
        try:
            keep_signed_in_label = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'label[for="input36"][data-se-for-name="rememberMe"]'))
            )
            keep_signed_in_label.click()
            time.sleep(5)
        except TimeoutException:
            print("No 'Keep me signed in' option found. Skipping.")
        try:
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.button.button-primary[type="submit"][value="Next"]'))
            )
            next_button.click()
            time.sleep(10)
        except TimeoutException:
            print("No 'Next' button found. Skipping.")
        print("🎉 Login successful! Reloading the page...")
        try:
            alert = driver.switch_to.alert
            alert.accept()
        except NoAlertPresentException:
            print("No 'Alert' button found. Skipping.")

        driver.get(expedia_url)
    except TimeoutException:
        print("✅ No sign-in required. Continuing with expedia_url...")

check_and_login(driver, expedia_url)

wait = WebDriverWait(driver, 10)

try:
    settings_buttons = wait.until(lambda d: d.find_elements(By.CSS_SELECTOR, "button.settingsButton"))
    if len(settings_buttons) >= 2:
        settings_buttons[0].click()
    else:
        raise TimeoutException("Không tìm thấy nút Settings thứ 2")

    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.uitk-menu-container[aria-hidden='false']")))

    download_csv_button = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//div[@class='uitk-menu-container uitk-menu-open uitk-menu-pos-left uitk-menu-container-autoposition'][@aria-hidden='false']//button[contains(@class, 'uitk-list-item')]//span[text()='Download CSV']/ancestor::button")
        )
    )
    download_csv_button.click()

except TimeoutException:
    pass

wait = WebDriverWait(driver, 10)
wait.until(EC.url_contains(path_fragment))
time.sleep(wait_seconds)

# %%
source_folder = r"C:\Users\huuchinh.nguyen\Downloads"
destination_folder = r"C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\Rawdata\CAPTURE\current_agent"

file_patterns = [os.path.join(source_folder, "Logged-In*.csv"),
                 os.path.join(source_folder, "Logged-In*.xlsx")]

for pattern in file_patterns:
    for filepath in glob.glob(pattern):
        filename = os.path.basename(filepath)
        destination_path = os.path.join(destination_folder, filename)
        shutil.move(filepath, destination_path)
        print(f"Đã chuyển: {filepath} -> {destination_path}")

# %%
from datetime import datetime, timedelta

def convert_to_datetime(struct_time):
    return datetime(*struct_time[:6])

def input_data(data_dir):
    list_files = []
    
    for filename in pathlib.Path(data_dir).glob('**/*.*'):
        file_suffixes = filename.suffixes
        if file_suffixes and file_suffixes[-1].lower() in ['.xlsx', '.csv']:
            export_time = os.path.getmtime(filename)
            export_time_datetime = convert_to_datetime(time.localtime(export_time))
            file_name = filename.stem
            
            try:
                if file_suffixes[-1].lower() == '.xlsx':
                    df = pl.read_excel(filename)
                    if df.is_empty():
                        continue
                    df = df.with_columns([
                        pl.lit(file_name).alias('sheet_name'),
                        pl.lit(export_time_datetime).alias('Export time')
                    ])
                    list_files.append(df)
                    
                elif file_suffixes[-1].lower() == '.csv':
                    if os.path.getsize(filename) == 0:
                        continue
                    df = pl.read_csv(filename)
                    if df.is_empty():
                        continue
                    df = df.with_columns([
                        pl.lit(file_name).alias('sheet_name'),
                        pl.lit(export_time_datetime).alias('Export time')
                    ])
                    list_files.append(df)
            except Exception as e:
                continue
                
    if list_files:
        return pl.concat(list_files)
    else:
        return pl.DataFrame()
    
outage_db = input_data('current_agent')
outage_db = outage_db.sort(["Export time"])
latest16 = (
    outage_db.select("Export time")
    .unique()
    .sort("Export time", descending=True)
    .limit(16)
)   
outage_db = outage_db.filter(pl.col("Export time").is_in(latest16.get_column("Export time")))

outage_db = outage_db.select(['Business Location','Export time', 'Agent Name','Agent Manager','Connect State','Assigned Workitem Count','Duration', 'Queue Group / Routing Profile'])
outage_db = outage_db.with_columns(
    pl.col("Export time").cast(pl.Datetime, strict=False).alias("VN_Datetime")
).with_columns([
    pl.col("VN_Datetime").dt.date().alias("VN_Date"),
    pl.col("VN_Datetime").dt.truncate("30m").dt.strftime("%H:%M").alias("VN_Intervals"),
])
outage_db = outage_db.with_columns(
    (pl.col("VN_Datetime") - pl.duration(hours=14)).alias("_pst_tmp")
).with_columns([
    pl.col("_pst_tmp").dt.date().alias("PST_Date"),
    pl.col("_pst_tmp").dt.truncate("30m").dt.strftime("%H:%M").alias("PST_Intervals"),
]).drop("_pst_tmp")
outage_db = outage_db.with_columns([
    pl.col("VN_Date").cast(pl.Datetime).alias("_vn_base"),
    pl.col("VN_Intervals").str.slice(0, 2).cast(pl.Int64).alias("_vn_h"),
    pl.col("VN_Intervals").str.slice(3, 2).cast(pl.Int64).alias("_vn_m"),
]).with_columns(
    (pl.col("_vn_base") + pl.duration(hours=pl.col("_vn_h"), minutes=pl.col("_vn_m"))).alias("VN_Datetime")
).drop(["_vn_base", "_vn_h", "_vn_m"])
outage_db = outage_db.with_columns([
    pl.col("PST_Date").cast(pl.Datetime).alias("_pst_base"),
    pl.col("PST_Intervals").str.slice(0, 2).cast(pl.Int64).alias("_pst_h"),
    pl.col("PST_Intervals").str.slice(3, 2).cast(pl.Int64).alias("_pst_m"),
]).with_columns(
    (pl.col("_pst_base") + pl.duration(hours=pl.col("_pst_h"), minutes=pl.col("_pst_m"))).alias("PST_Datetime")
).drop(["_pst_base", "_pst_h", "_pst_m"])
outage_db = outage_db.with_columns(pl.col('Duration').str.strptime(pl.Time, format='%H:%M:%S'))
outage_db = (outage_db.with_columns(
    pl.when(pl.col("Queue Group / Routing Profile").is_in(["Chat_OD_EN_Car_Activity", "Chat_OD_EN_Lodging", "Chat - Global English Lodging Nesting", "Chat_Lodging English w Car"])).then(pl.lit("Lodging Chat"))
    .when(pl.col("Queue Group / Routing Profile").is_in(["Chat - Global English Non- Lodging Nesting", "Chat_OD_EN_Dual_GDS"])).then(pl.lit("Non Lodging Chat")).otherwise(None).alias("LOB")))
hcm_outage = outage_db.filter(pl.col("Business Location") == "Concentrix (Ho Chi Minh City)")

duration_seconds = (
    pl.col("Duration").dt.hour().cast(pl.Int32) * 3600 +
    pl.col("Duration").dt.minute().cast(pl.Int32) * 60 +
    pl.col("Duration").dt.second().cast(pl.Int32)
)
hcm_outage = hcm_outage.with_columns(duration_seconds.alias("duration_seconds"))
break_lunch_outage = hcm_outage.filter(
    pl.col("Connect State").is_in(["BREAK", "LUNCH"])
).with_columns(
    pl.when((pl.col("Connect State") == "BREAK") & (pl.col("Assigned Workitem Count").is_null()) &(pl.col("duration_seconds") > 15 * 60))
    .then(pl.lit("over-break"))
    .when((pl.col("Connect State") == "LUNCH") & (pl.col("Assigned Workitem Count").is_null()) & (pl.col("duration_seconds") > 60 * 60))
    .then(pl.lit("over-lunch"))
    .otherwise(None)
    .alias("Note")
)
coaching_training_outage = hcm_outage.filter(
    pl.col("Connect State").is_in(["COACHING", "TRAINING"])
).with_columns(
    pl.lit("need to check").alias("Note")
)
active_outage = hcm_outage.filter(
    pl.col("Connect State").is_in(["AVAILABLECHAT", "OUTBOUNDCALL"])
).with_columns(
    pl.when((pl.col("Connect State") == "AVAILABLECHAT") & (pl.col("Assigned Workitem Count").is_null()) & (pl.col("duration_seconds") > 30 * 60))
    .then(pl.lit("need to check"))
    .when((pl.col("Connect State") == "OFFLINE") & (pl.col("Assigned Workitem Count") >= 1) & (pl.col("duration_seconds") > 10 * 60))
    .then(pl.lit("need to check"))
    .otherwise(None)
    .alias("Note")
)
hcm_outage_dup = outage_db.with_columns(
    (
        pl.when(pl.col("Assigned Workitem Count").cast(pl.Int64, strict=False).fill_null(0) >= 1).then(pl.lit("Available"))
        .when((pl.col("Connect State") == "AVAILABLECHAT") | (pl.col("Connect State") == "OUTBOUNDCALL") | (pl.col("Connect State") == "READY")).then(pl.lit("Available"))
        .when(pl.col("Connect State").str.contains("BREAK", literal=True)).then(pl.lit("Break"))
        .when(pl.col("Connect State").str.contains("LUNCH", literal=True)).then(pl.lit("Lunch"))
        .when(pl.col("Connect State").str.contains("COACHING", literal=True)).then(pl.lit("Coaching"))
        .when(pl.col("Connect State").str.contains("TRAINING", literal=True)).then(pl.lit("Training"))
        .when(pl.col("Connect State").str.contains("TEAMMEETING", literal=True)).then(pl.lit("Team Meeting"))
        .otherwise(pl.lit("Other"))
    ).alias("Category"),
    pl.lit("all").alias("all")
)
grouped = (
    hcm_outage_dup
    .group_by([
        "VN_Date","PST_Date","VN_Datetime","PST_Datetime",
        "VN_Intervals","PST_Intervals","Business Location","LOB","Category"
    ])
    .agg(pl.col("Agent Name").n_unique().alias("Count"))
    .sort("VN_Datetime", descending=True)
)
pivot_hcm = (
    grouped
    .filter(pl.col("Business Location").str.contains("Ho Chi Minh",literal=False))
    .pivot(
        values="Count",
        index=[
            "VN_Date", "PST_Date", "VN_Datetime", "PST_Datetime",
            "VN_Intervals", "PST_Intervals", "Business Location", "LOB"
        ],
        columns="Category"
    )
    .fill_null(0)
)
if "Available" not in pivot_hcm.columns:
    if "Available" in pivot_hcm.columns and "Outbound" in pivot_hcm.columns:
        pivot_hcm = pivot_hcm.with_columns(
            (pl.col("Available") + pl.col("Outbound")).alias("Available")
        ).drop(["Available", "Outbound"])
    elif "Available" in pivot_hcm.columns:
        pivot_hcm = pivot_hcm.rename({"Available": "Available"})
    elif "Outbound" in pivot_hcm.columns:
        pivot_hcm = pivot_hcm.rename({"Outbound": "Available"})
    else:
        pivot_hcm = pivot_hcm.with_columns(pl.lit(0).alias("Available"))

for c in ["Break", "Lunch", "Coaching", "Training", "Other"]:
    if c not in pivot_hcm.columns:
        pivot_hcm = pivot_hcm.with_columns(pl.lit(0).alias(c))
pivot_hcm = (
    OU_Raw.select([
        "PST_Date", "VN_Date", "PST_Datetime", "VN_Datetime", "PST_Intervals", "VN_Intervals", "value",
        pl.col("value").alias("HC Required"),
    ])
    .join(
        pivot_hcm.select(["VN_Datetime","Available","Break","Lunch","Coaching","Training","Other"]),
        on="VN_Datetime",
        how="left",
    ).join(
        Schedule_Raw.select(["VN_Datetime","Scheduled", "UPL", "PL", "Meal|Break"]),
        on="VN_Datetime",
        how="left",
    )
)
pivot_hcm = pivot_hcm.with_columns([
    pl.col("VN_Date").dt.strftime("%b %d").alias("VN_Date"),
    pl.col("PST_Date").dt.strftime("%b %d").alias("PST_Date"),
])
pivot_hcm = pivot_hcm.with_columns([
    (pl.col("Available") - pl.col("HC Required")).alias("Surplus/ Deficit")
]).with_columns([
    pl.col("HC Required").round(2),
    pl.col("Surplus/ Deficit").round(2)
])
def filter_around_current(pivot_df: pl.DataFrame, k: int = 8) -> pl.DataFrame:
    now = datetime.now()
    center = now.replace(minute=(now.minute // 30) * 30, second=0, microsecond=0)

    start = center - timedelta(minutes=30 * k)
    end   = center + timedelta(minutes=30 * k)

    return pivot_df.filter(pl.col("VN_Datetime").is_between(start, end, closed="both"))
pivot_hcm = filter_around_current(pivot_hcm, k=8)
select_cols = ["VN_Date","VN_Intervals","PST_Date","PST_Intervals","Scheduled", "UPL", "PL", "Meal|Break","HC Required", "Surplus/ Deficit", "Available", "Break", "Lunch", "Coaching", "Training", "Other"]
select_cols = [c for c in select_cols if c in pivot_hcm.columns]
pivot_hcm = pivot_hcm.select(select_cols)
pivot_staffing_monitoring = pivot_hcm.to_pandas().fillna("")
pivot_staffing_monitoring

# %%
latest_export = outage_db.select(pl.col("Export time").cast(pl.Datetime, strict=False).max()).item()
last_refresh_txt = latest_export.strftime("%b %d, %Y %H:%M") if latest_export is not None else "N/A"
print("Last refresh (VN):", last_refresh_txt)

# %%
import matplotlib.pyplot as plt
from pandas.plotting import table
from PIL import Image
import io
import datetime

def df_to_image(df, title, is_global=False):
    now = datetime.datetime.now()
    if is_global:
        now = now - datetime.timedelta(hours=14)
    
    realtime = now.strftime("%I:%M %p")
    realtime_date = now.strftime("%d-%b-%Y")
    timezone = "(PST)" if is_global else "(VNT)"

    df_no_index = df.reset_index(drop=True)
    fig, ax = plt.subplots(figsize=(10, 8))
    ax.axis('off')
    ax.set_title(f'{title} updated on {realtime_date} at {realtime} {timezone}', fontsize=16, fontweight='bold', loc='center', pad=1)
    
    if not df_no_index.empty:
        col_widths = [max(df_no_index[col].apply(lambda x: len(str(x)))) for col in df_no_index.columns]

    tbl = table(ax, df_no_index, loc='center', cellLoc='center')
    for i, col in enumerate(df_no_index.columns):
        tbl.auto_set_column_width([i])

    tbl.auto_set_font_size(False)
    tbl.set_fontsize(10)
    tbl.scale(1.2, 1.2)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=300)
    buf.seek(0)

    image = Image.open(buf)
    # plt.close(fig)
    return image

def save_df_image_if_not_empty(df, title, is_global=False):
    if df.empty:
        print(f"Warning: DataFrame '{title}' is empty, skipping image generation.")
        return None
    else:
        return df_to_image(df, title, is_global=is_global)

img_pivot_hcm = save_df_image_if_not_empty(pivot_staffing_monitoring, "Realtime Staffing Monitoring", is_global=False)

# %%
def send_report_to_group(driver, teams_url, group_chat_name, report_name, image, LC_Cases, is_global=False):
    if LC_Cases == 0:
        print(f"Skipping sending report '{report_name}' because LC_Cases == 0")
        return
    
    send_to_clipboard(image)
    current_url = driver.current_url
    
    if not current_url.startswith(teams_url):
        driver.get(teams_url)
        time.sleep(8) 
        
    wait = WebDriverWait(driver, 20)
    clicked_success = False
    attempts = 0
    
    while attempts < 3 and not clicked_success:
        try:
            print(f"Searching for group '{group_chat_name}' (Attempt {attempts + 1})...")
            group_chat = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[contains(text(),'{group_chat_name}')]")))      
            group_chat.click()
            clicked_success = True
            print(f"✅ Successfully entered group '{group_chat_name}'")
        except StaleElementReferenceException:
            print("⚠️ Stale Element detected (DOM changed), retrying...")
            time.sleep(2)
            attempts += 1
        except Exception as e:
            print(f"❌ Error clicking chat group: {e}")
            return

    if not clicked_success:
        print(f"❌ Failed to click on group '{group_chat_name}' after 3 attempts.")
        return
        
    try:
        time.sleep(3) 
        chat_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[role='textbox']")))
        chat_box.click()
        
        now = datetime.datetime.now()
        if is_global:
            now = now - datetime.timedelta(hours=14)
        realtime = now.strftime("%I:%M %p")  
        timezone = "(PST)" if is_global else "(VNT)"
        chat_text = f"The {report_name} was updated at {realtime} {timezone}"
        
        actions = ActionChains(driver)
        actions.send_keys("@everyone").perform()
        time.sleep(1)
        actions.send_keys(Keys.ENTER).perform()
        actions.send_keys(" ")
        actions.send_keys(chat_text)
        actions.perform()
        time.sleep(1)
        
        actions = ActionChains(driver)
        actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        print("Waiting for image to upload to chat box...")
        time.sleep(5) 
        
        actions.send_keys(Keys.ENTER).perform()
        print("✅ Report sent successfully!")
        time.sleep(5)
        
    except Exception as e:
        print(f"❌ Error during input/sending message: {e}")

# Call function
send_report_to_group(driver, teams_url, "EXPEDIA - KPI Update", "Realtime Staffing Monitoring", img_pivot_hcm, 2, is_global=False)

driver.quit()

# %%
# from openpyxl import Workbook, load_workbook
# from datetime import datetime

# end_time = datetime.now()
# execution_time = (end_time - start_time).total_seconds()

# log_file = "bot_log/bot_capture_log.xlsx"

# if not os.path.exists(log_file):
#     wb = Workbook()
#     ws = wb.active
#     ws.title = "Log"
#     ws.append(["BOT Name", "Run at", "End at", "Execution Time (seconds)"])
# else:
#     wb = load_workbook(log_file)
#     ws = wb["Log"]

# ws.append([
#     report_name,
#     start_time.strftime('%Y-%m-%d %H:%M:%S'),
#     end_time.strftime('%Y-%m-%d %H:%M:%S'),
#     execution_time
# ])

# wb.save(log_file)


