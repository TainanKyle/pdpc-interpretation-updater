@echo off
REM Build the Windows .exe with PyInstaller.
REM Run on Windows (PyInstaller cannot cross-build for other OSes).
cd /d "%~dp0"

python -m venv venv
call venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt pyinstaller

REM --collect-submodules truststore: its backend is chosen behind a platform
REM test, so PyInstaller's static analysis cannot see the Windows submodule.
pyinstaller --name "PDPC函釋更新工具" --windowed --onefile --noconfirm --collect-submodules truststore app.py

echo.
echo Done. The exe is at: dist\PDPC函釋更新工具.exe
pause
