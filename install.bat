@echo off
title Humanforce - Google Calendar Sync Setup
color 0A

echo ====================================================
echo   Humanforce - Google Calendar Sync - One Click Setup
echo ====================================================
echo.

REM --- Check for Python ---
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo [ERROR] Python is not installed on your system.
    echo Please install Python 3.10+ from https://www.python.org/downloads/
    pause
    start https://www.python.org/downloads/
    exit /b
) ELSE (
    for /f "tokens=2 delims= " %%a in ('python --version') do set PYVER=%%a
    echo Found Python version: %PYVER%
)

echo.
echo [1/3] Installing Python dependencies...
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo [2/3] Installing Playwright browser (Chromium)...
python -m playwright install chromium

echo.
echo [3/3] Setup complete! You can now run the app with:
echo python app.py
echo.

pause
