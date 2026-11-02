@echo off
setlocal
REM Run from this script's folder so relative paths resolve consistently.
cd /d "%~dp0"

REM Check for the Windows Python launcher before attempting setup.
where py >nul 2>nul
if errorlevel 1 (
  echo Python was not found.
  echo Install Python 3.12 from python.org and select "Add Python to PATH".
  pause
  exit /b 1
)

REM First launch creates an isolated Python environment and installs dependencies.
if not exist ".venv\Scripts\python.exe" (
  echo Setting up the roster updater for the first time...
  py -3 -m venv .venv
  if errorlevel 1 goto :failed
  ".venv\Scripts\python.exe" -m pip install --upgrade pip
  if errorlevel 1 goto :failed
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto :failed
)

REM pythonw starts the GUI without keeping a console window open.
start "" ".venv\Scripts\pythonw.exe" roster_app.py
exit /b 0

:failed
echo.
echo Setup did not finish. Send the error shown above to the person supporting the app.
pause
exit /b 1

