"""每次執行後自動將程式碼推送到 GitHub"""
import subprocess, shutil, os

REPO = r"C:\Users\gjsky\OneDrive\桌面\磐鈺雲華商業停車場"
SRC  = r"C:\temp\parking"

FILES_TO_SYNC = [
    (os.path.join(SRC, "import.py"),    os.path.join(REPO, "主程式", "import.py")),
    (os.path.join(SRC, "catch_up.py"),  os.path.join(REPO, "主程式", "catch_up.py")),
    (os.path.join(SRC, "run_daily.bat"),os.path.join(REPO, "主程式", "run_daily.bat")),
    (os.path.join(SRC, "對話記錄.md"), os.path.join(REPO, "主程式", "對話記錄.md")),
]

def run(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", cwd=REPO, **kw)

# 同步檔案
for src, dst in FILES_TO_SYNC:
    if os.path.exists(src):
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)

# git add
files = [os.path.relpath(dst, REPO) for _, dst in FILES_TO_SYNC]
files += [
    os.path.join("每日現金支付統計", "import.py"),
    os.path.join("每日現金支付統計", "匯入停車資料.py"),
    os.path.join("每日現金支付統計", "每日匯入.bat"),
    ".gitignore",
]
run(["git", "add"] + files)

# 有變動才 commit + push
diff = run(["git", "diff", "--cached", "--quiet"])
if diff.returncode != 0:
    from datetime import datetime
    msg = f"自動更新 {datetime.now().strftime('%Y-%m-%d')}"
    run(["git", "commit", "-m", msg])
    result = run(["git", "push"])
    if result.returncode == 0:
        print("  [GitHub] 推送成功")
    else:
        print(f"  [GitHub] 推送失敗：{result.stderr.strip()}")
else:
    print("  [GitHub] 無變動，略過推送")
