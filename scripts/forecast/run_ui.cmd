@echo off
setlocal
REM Start the Forecast web UI and open it in the browser.

cd /d "%~dp0..\.."
set "ROOT=%CD%"
if not defined FORECAST_UI_PORT set "FORECAST_UI_PORT=8765"
set "URL=http://127.0.0.1:%FORECAST_UI_PORT%"

if exist "%ROOT%\.venv\Scripts\python.exe" (
  set "PYTHON=%ROOT%\.venv\Scripts\python.exe"
) else (
  set "PYTHON=python"
)

"%PYTHON%" -c "import fastapi, uvicorn" >nul 2>&1
if errorlevel 1 (
  echo Installing Forecast UI dependencies...
  "%PYTHON%" -m pip install --index-url https://pypi.org/simple -r "%ROOT%\scripts\forecast\requirements.txt"
  if errorlevel 1 exit /b 1
)

echo Forecast UI -^> %URL%
start "" "%URL%"
"%PYTHON%" -m uvicorn ui.app:app --app-dir "%ROOT%\scripts\forecast" --host 127.0.0.1 --port %FORECAST_UI_PORT%
