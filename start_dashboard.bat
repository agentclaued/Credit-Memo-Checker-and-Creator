@echo off
rem Double-click to run: checks Python, installs requirements on first run, starts the server.
cd /d "%~dp0"
python --version >nul 2>&1 || (echo Python is not installed. & pause & exit /b 1)
python -c "import fastapi, uvicorn, jinja2, multipart, openpyxl, python_calamine, pandas" >nul 2>&1 || python -m pip install -r requirements.txt
for /f "delims=" %%H in ('python -c "import config;print(config.HOST)"') do set HOST=%%H
for /f "delims=" %%P in ('python -c "import config;print(config.PORT)"') do set PORT=%%P
echo Starting on http://%HOST%:%PORT%  (Ctrl+C to stop)
python -m uvicorn app.main:app --host %HOST% --port %PORT% --reload
pause
