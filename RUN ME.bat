@echo off
title TRU Funding Client Analyzer
color 1F

echo.
echo  =======================================
echo    TRU Funding Client Analyzer
echo  =======================================
echo.

:: Check if Python is installed
py --version >nul 2>&1
if errorlevel 1 (
    echo  Python not found. Opening download page...
    echo  Please install Python from python.org
    echo  IMPORTANT: Check "Add Python to PATH" during install!
    start https://www.python.org/downloads/
    pause
    exit
)

echo  Installing required packages...
py -m pip install flask pdfplumber openpyxl --quiet --disable-pip-version-check

echo  Starting tool...
echo.
echo  ----------------------------------------
echo   Open your browser to: http://localhost:5050
echo  ----------------------------------------
echo.

:: Open browser automatically after 2 seconds
start "" timeout /t 2 >nul
start "" "http://localhost:5050"

py app.py

pause
