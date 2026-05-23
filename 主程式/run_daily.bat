@echo off
cd /d C:\temp\parking
"C:\Users\gjsky\AppData\Local\Programs\Python\Python312\python.exe" catch_up.py >> C:\temp\parking\run_log.txt 2>&1
"C:\Users\gjsky\AppData\Local\Programs\Python\Python312\python.exe" check_daily.py >> C:\temp\parking\run_log.txt 2>&1
"C:\Users\gjsky\AppData\Local\Programs\Python\Python312\python.exe" sync_github.py >> C:\temp\parking\run_log.txt 2>&1
