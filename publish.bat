@echo off
rem ============================================================
rem  EPC Dashboard - Build + Publish to Netlify
rem  Double-click: builds from the ERP export, then publishes
rem  (or opens Netlify Drop if the CLI is not installed)
rem ============================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found. Install from https://python.org (check "Add to PATH").
    pause
    exit /b 1
)

echo [1/2] Building dashboard...
python tools\update_dashboard.py %*
if errorlevel 1 (
    echo [FAILED] Build failed - see above.
    pause
    exit /b 1
)

where netlify >nul 2>nul
if errorlevel 1 (
    echo.
    echo [2/2] Netlify CLI not installed.
    echo.
    echo   Opening Netlify in your browser...
    echo   Drag the  dist  folder onto the page.
    echo   First time: creates the site. Next times: updates it.
    start "" "https://app.netlify.com/drop"
    start "" "%~dp0dist"
    pause
    exit /b 0
)

echo [2/2] Publishing to Netlify...
netlify deploy --prod --dir=dist
if errorlevel 1 (
    echo [FAILED] Deploy failed. First time? run:  netlify login   then   netlify link
    pause
    exit /b 1
)
echo.
echo Published! Your link is live.
timeout /t 5 >nul
