@echo off
chcp 65001 >nul 2>&1
title Customer Price Manager v5.1

echo.
echo Starting Customer Price Manager...
echo.

REM Install packages if needed (first run only)
python -m pip install -q streamlit pandas openpyxl 2>nul

REM Run the app
python -m streamlit run app.py

pause
