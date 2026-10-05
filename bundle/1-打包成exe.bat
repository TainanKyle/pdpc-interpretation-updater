@echo off
rem Build the exe locally so it carries no Mark-of-the-Web.
rem Pure-ASCII commands; Chinese appears only in echo text (after chcp 65001).
chcp 65001 >nul
cd /d "%~dp0"
title PDPC hanshi tool - build exe

if not exist "WinPython\python.exe" (
    echo [錯誤] 找不到 WinPython\python.exe
    echo 請確認整個 zip 已「全部解壓縮」後再執行（不要在壓縮檔內直接雙擊）。
    goto :fail
)

echo === 步驟 1/3：安裝套件（離線，約 1 分鐘）===
"WinPython\python.exe" -m pip install --no-index --find-links=wheels --disable-pip-version-check -q requests==2.34.2 openpyxl==3.1.5 pyinstaller==6.21.0
if errorlevel 1 goto :fail

rem truststore is OPTIONAL (OS certificate store support). Installed separately
rem and NEVER fatal: an older extracted bundle has no truststore wheel, and
rem _build_exe.py only passes --collect-submodules when it is actually present.
"WinPython\python.exe" -m pip install --no-index --find-links=wheels --disable-pip-version-check -q truststore==0.10.4 2>nul

echo === 步驟 2/3：打包執行檔（約 2-5 分鐘，請耐心等候）===
"WinPython\python.exe" _build_exe.py
if errorlevel 1 goto :fail

echo === 步驟 3/3：整理輸出 ===
move /y "dist\*.exe" . >nul
if errorlevel 1 goto :fail
rmdir /s /q build dist 2>nul
del /q *.spec 2>nul

echo.
echo ==============================================
echo   完成！執行檔已產生在本資料夾：
echo   PDPC函釋更新工具.exe
echo   之後直接雙擊它即可使用，不必再跑本程式。
echo ==============================================
pause
exit /b 0

:fail
echo.
echo ==============================================
echo   發生錯誤，打包未完成。
echo   請把上面的訊息截圖，到 GitHub 開 issue 回報：
echo   https://github.com/TainanKyle/pdpc-interpretation-updater/issues
echo   （替代方案：改雙擊「2-直接執行.bat」，
echo     不需打包也能使用。）
echo ==============================================
pause
exit /b 1
