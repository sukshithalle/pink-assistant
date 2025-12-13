@echo off
chcp 65001 >nul
echo ======================================
echo  Pink Assistant — One click installer
echo ======================================
echo.

REM Ensure python is on PATH
where python >nul 2>&1
IF ERRORLEVEL 1 (
  echo Python not found on PATH. Install Python 3.9+ and ensure 'python' is in PATH.
  pause
  exit /b 1
)

REM Upgrade pip and install required packages (best-effort)
echo Installing / checking required Python packages...
python -m pip install --upgrade pip setuptools wheel
python -m pip install pyttsx3 SpeechRecognition psutil screen-brightness-control pywin32 pyautogui pyaudio winsound
python -m pip install speechrecognition pyttsx3 psutil pyautogui
python -m pip install screen-brightness-control pywin32 pygetwindow
python -m pip install opencv-python mediapipe numpy

REM Note: pyaudio install may fail on Windows without wheels; if it fails, follow instructions at:
REM https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio

echo.
echo Launching Pink Assistant...
python main.py
pause
