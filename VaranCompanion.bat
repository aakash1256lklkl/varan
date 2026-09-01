@echo off
REM Varan Companion launcher - open the floating AI panel beside Office.
cd /d "%~dp0"
start "" pythonw companion_run.py
echo Varan Companion started. Press Ctrl+Alt+Shift+V to show/hide it.
timeout /t 3 >nul
