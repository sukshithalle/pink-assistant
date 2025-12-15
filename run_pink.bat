@echo off
chcp 65001 >nul
echo ======================================
echo  Pink Assistant — Starting
echo ======================================

REM Ensure Python 3.10 exists
py -3.10 --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo Python 3.10 not found. Please install Python 3.10.
    pause
    exit /b 1
)

REM Create venv if not exists
if not exist venv310 (
    echo Creating virtual environment...
    py -3.10 -m venv venv310
)

REM Activate venv
call venv310\Scripts\activate

REM Install dependencies
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo Launching Pink Assistant...
python main.py
pause
