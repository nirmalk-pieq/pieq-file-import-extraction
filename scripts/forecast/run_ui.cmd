@echo off
setlocal
REM Start the Forecast web UI and open it in the browser.

cd /d "%~dp0..\.."
set "ROOT=%CD%"
if not defined FORECAST_UI_PORT set "FORECAST_UI_PORT=8765"
set "URL=http://127.0.0.1:%FORECAST_UI_PORT%"

if exist "%ROOT%\.venv\Scripts\python.exe" (
  set "PYTHON=%ROOT%\.venv\Scripts\python.exe"
) else if exist "%ROOT%\.venv\Scripts\python" (
  set "PYTHON=%ROOT%\.venv\Scripts\python"
) else (
  where python >nul 2>&1
  if errorlevel 1 (
    echo Python not found. Create a venv or install Python and retry.
    exit /b 1
  )
  set "PYTHON=python"
)

"%PYTHON%" -c "import fastapi, uvicorn, multipart" >nul 2>&1
if errorlevel 1 (
  echo Installing Forecast UI dependencies...
  "%PYTHON%" -m pip install --index-url https://pypi.org/simple -r "%ROOT%\scripts\forecast\requirements.txt"
  if errorlevel 1 exit /b 1
)

echo Forecast UI -^> %URL%
REM Give the server a moment, then open the browser.
start "" cmd /c "timeout /t 1 /nobreak >nul & start \"\" \"%URL%\""
"%PYTHON%" -m uvicorn ui.app:app --app-dir "%ROOT%\scripts\forecast" --host 127.0.0.1 --port %FORECAST_UI_PORT%
