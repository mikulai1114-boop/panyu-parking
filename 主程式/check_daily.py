"""每日確認前一天資料是否已更新，缺漏時自動補跑"""
import os, glob
from datetime import datetime, timedelta
import openpyxl

CONFIG = r"C:\temp\parking\config.txt"

with open(CONFIG, encoding="utf-8") as f:
    WORKBOOK = f.read().strip()

MONTHLY_DIR = os.path.join(os.path.dirname(WORKBOOK), "月份資料")
MONTHS_NUM  = ["01","02","03","04","05","06","07","08","09","10","11","12"]

def _monthly_path(date_obj):
    return os.path.join(MONTHLY_DIR, f"{date_obj.year}年{MONTHS_NUM[date_obj.month-1]}月.xlsx")

def check(date_obj):
    """檢查指定日期的資料是否存在，回傳 (ok, total, records)"""
    sheet_name = date_obj.strftime("%Y-%m-%d")
    mpath = _monthly_path(date_obj)

    if not os.path.exists(mpath):
        return False, 0, 0

    try:
        wb = openpyxl.load_workbook(mpath, data_only=True)
    except Exception:
        return False, 0, 0

    if sheet_name not in wb.sheetnames:
        return False, 0, 0

    ws = wb[sheet_name]
    # 找合計列（B欄="合計"）
    total, records = 0, 0
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "合計":
            v = ws.cell(r, 4).value
            if isinstance(v, (int, float)):
                total = int(v)
            break
        if ws.cell(r, 1).value is not None:
            records += 1

    return True, total, max(records - 1, 0)  # 扣掉表頭


def main():
    import importlib.util
    yesterday = (datetime.now() - timedelta(days=1)).date()
    date_str  = yesterday.strftime("%Y-%m-%d")

    ok, total, records = check(yesterday)

    if ok:
        print(f"  [每日確認] {date_str} OK  {records} 筆  合計 ${total:,}")
    else:
        print(f"  [每日確認] {date_str} 資料缺漏，自動補跑...")
        spec = importlib.util.spec_from_file_location("parking", r"C:\temp\parking\import.py")
        mod  = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        try:
            mod.run(date_str)
            print(f"  [每日確認] 補跑完成")
        except Exception as e:
            print(f"  [每日確認] 補跑失敗：{e}")

if __name__ == "__main__":
    main()
