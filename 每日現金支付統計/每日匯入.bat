@echo off
chcp 65001 > nul
echo ========================================
echo   磐鈺雲華商業停車場 每日資料匯入
echo ========================================
echo.
set "PYPATH=C:\Users\gjsky\AppData\Local\Programs\Python\Python312\python.exe"
set "SCRIPT=C:\Users\gjsky\OneDrive\桌面\磐鈺雲華商業停車場\每日現金支付統計\匯入停車資料.py"

if "%1"=="" (
    echo 匯入昨日資料...
    "%PYPATH%" "%SCRIPT%"
) else (
    echo 匯入 %1 資料...
    "%PYPATH%" "%SCRIPT%" "%1"
)

echo.
if %errorlevel%==0 (
    echo [成功] 資料已匯入完成！
) else (
    echo [錯誤] 請檢查網路連線或後台帳號
)
pause