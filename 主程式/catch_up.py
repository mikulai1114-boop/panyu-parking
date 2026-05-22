"""每日啟動時自動偵測缺漏日期並逐日補跑"""
import os, sys, importlib.util
from datetime import datetime, timedelta

LAST_RUN_FILE = r"C:\temp\parking\last_run.txt"

# 載入 import.py
spec = importlib.util.spec_from_file_location("parking", r"C:\temp\parking\import.py")
mod  = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

today     = datetime.now().date()
yesterday = today - timedelta(days=1)

# 讀取上次執行日期
if os.path.exists(LAST_RUN_FILE):
    with open(LAST_RUN_FILE, encoding="utf-8-sig") as f:
        last_str = f.read().strip()
    try:
        last_run = datetime.strptime(last_str, "%Y-%m-%d").date()
    except ValueError:
        last_run = yesterday
else:
    last_run = yesterday

# 計算需補跑的日期（上次執行日的隔天 ~ 昨天）
start = last_run + timedelta(days=1)
dates = []
d = start
while d <= yesterday:
    dates.append(d)
    d += timedelta(days=1)

if not dates:
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] 無需補跑")
else:
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] 需補跑 {len(dates)} 天：{dates[0]} ~ {dates[-1]}")

for d in dates:
    date_str = d.strftime("%Y-%m-%d")
    print(f"\n===== 補跑 {date_str} =====")
    try:
        mod.run(date_str)
        # 更新上次執行日期
        with open(LAST_RUN_FILE, "w", encoding="utf-8") as f:
            f.write(date_str)
    except Exception as e:
        print(f"  補跑失敗：{e}")

# 全部跑完後更新記錄
if dates:
    with open(LAST_RUN_FILE, "w", encoding="utf-8") as f:
        f.write(yesterday.strftime("%Y-%m-%d"))
    print(f"\n全部補跑完成！")
elif not os.path.exists(LAST_RUN_FILE):
    with open(LAST_RUN_FILE, "w", encoding="utf-8") as f:
        f.write(yesterday.strftime("%Y-%m-%d"))
