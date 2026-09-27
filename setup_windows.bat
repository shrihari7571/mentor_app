@echo off
cd /d "%~dp0"
echo Installing required Python packages...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Installation failed. Check that Python is installed and available in PATH.
  pause
  exit /b 1
)
echo.
echo Starting Mentor-Mentee Tracker...
python -m streamlit run app.py
pause
