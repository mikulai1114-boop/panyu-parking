@echo off
cd /d C:\temp\parking
"C:\Users\gjsky\AppData\Local\Programs\Python\Python312\python.exe" catch_up.py >> C:\temp\parking\run_log.txt 2>&1

:: 自動更新 GitHub
cd /d "C:\Users\gjsky\OneDrive\桌面\磐鈺雲華商業停車場"
:: 同步對話記錄
copy /Y "C:\temp\parking\對話記錄.md" "主程式\對話記錄.md" >> C:\temp\parking\run_log.txt 2>&1
git add 主程式\import.py 主程式\catch_up.py 主程式\run_daily.bat 主程式\安裝說明.md 主程式\對話記錄.md 每日現金支付統計\import.py 每日現金支付統計\匯入停車資料.py 每日現金支付統計\每日匯入.bat .gitignore >> C:\temp\parking\run_log.txt 2>&1
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "自動更新 %date%" >> C:\temp\parking\run_log.txt 2>&1
    git push >> C:\temp\parking\run_log.txt 2>&1
)
