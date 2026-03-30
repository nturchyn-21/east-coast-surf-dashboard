@echo off
title East Coast Surf Dashboard
cd /d "%~dp0"

echo ================================================
echo   Surf Dashboard - East Coast
echo ================================================
echo.

echo Checking dependencies...
pip install -r requirements.txt -q

echo.
echo Launching dashboard at http://localhost:8501
echo (Close this window to stop the dashboard)
echo.

python -m streamlit run surf_dashboard.py ^
    --server.headless false ^
    --browser.gatherUsageStats false ^
    --theme.primaryColor "#1F7A8C" ^
    --theme.backgroundColor "#f0f4f8" ^
    --theme.secondaryBackgroundColor "#ffffff" ^
    --theme.textColor "#1B3A5C"

pause
