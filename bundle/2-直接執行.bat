@echo off
rem Fallback entry: run app.py directly on the portable Python -- no exe is
rem built, so this works even where locally-built exes are blocked (AppLocker).
rem The console window stays open on purpose: closing it kills the app, and
rem it shows the error if anything goes wrong.
chcp 65001 >nul
cd /d "%~dp0"
title PDPC hanshi tool

if not exist "WinPython\python.exe" (
    echo [錯誤] 找不到 WinPython\python.exe
    echo 請確認整個 zip 已「全部解壓縮」後再執行（不要在壓縮檔內直接雙擊）。
    goto :fail
)

rem Install the two runtime packages offline on first run; skip when present.
"WinPython\python.exe" -c "import requests, openpyxl" 2>nul
if errorlevel 1 (
    echo === 首次執行：安裝套件（離線，約 1 分鐘）===
    "WinPython\python.exe" -m pip install --no-index --find-links=wheels --disable-pip-version-check -q requests==2.34.2 openpyxl==3.1.5
    if errorlevel 1 goto :fail
)

rem truststore is OPTIONAL (OS certificate store support; core.py runs without
rem it). Installed separately and NEVER fatal -- an older extracted bundle has
rem no truststore wheel, and a missing optional package must not block startup.
"WinPython\python.exe" -c "import truststore" 2>nul
if errorlevel 1 (
    "WinPython\python.exe" -m pip install --no-index --find-links=wheels --disable-pip-version-check -q truststore==0.10.4 2>nul
)

echo 視窗啟動中…（這個黑色視窗請不要關，關掉程式就結束了）
"WinPython\python.exe" app.py
if errorlevel 1 goto :fail
exit /b 0

:fail
echo.
echo 發生錯誤。請把上面的訊息截圖，到 GitHub 開 issue 回報：
echo https://github.com/TainanKyle/pdpc-interpretation-updater/issues
pause
exit /b 1
