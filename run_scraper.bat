@echo off
echo ===================================
echo Advanced Email Web Scraper
echo ===================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b
)

REM Check if requirements are installed
echo Checking and installing required packages...
pip install -r requirements.txt

echo.
echo Starting the scraper...
echo.

python local_scraper.py