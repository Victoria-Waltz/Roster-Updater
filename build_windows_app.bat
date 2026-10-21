@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.12 first.
  pause
  exit /b 1
)

if not exist ".buildvenv\Scripts\python.exe" (
  py -3 -m venv .buildvenv
  if errorlevel 1 goto :failed
)

".buildvenv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :failed
".buildvenv\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto :failed

".buildvenv\Scripts\pyinstaller.exe" --noconfirm --clean --onefile --windowed --name "TAIT Roster Updater" roster_app.py
if errorlevel 1 goto :failed

echo.
echo Finished. The application is in:
echo %CD%\dist\TAIT Roster Updater.exe
pause
exit /b 0

:failed
echo.
echo The Windows application could not be built. Review the error above.
pause
exit /b 1

