@echo off
rem ============================================================
rem  EPC Production Dashboard - Daily Build (double-click me)
rem ============================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Install it from https://python.org and check "Add Python to PATH".
    pause
    exit /b 1
)

python tools\update_dashboard.py %*

if errorlevel 1 (
    echo.
    echo [FAILED] See the error message above.
    pause
    exit /b 1
)

echo.
echo Opening standalone dashboard...
start "" "%~dp0dist\standalone.html"
timeout /t 4 >nul
