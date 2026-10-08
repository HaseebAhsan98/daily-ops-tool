@echo off
title Popular Group - Daily Ops Tool
echo ===================================================
echo     Popular Group - Daily Ops Tool Starting...
echo ===================================================
echo.
cd /d "%~dp0"
python -m streamlit run app.py
pause
