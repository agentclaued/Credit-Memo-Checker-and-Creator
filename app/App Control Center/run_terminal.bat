@echo off
rem Runs the server visibly in a terminal WITHOUT --reload (stops the background app first).
call "%~dp0stop_background.bat"
pushd "%~dp0..\.."
for /f "delims=" %%P in ('python -c "import config; print(config.PORT)"') do set PORT=%%P
for /f "delims=" %%H in ('python -c "import config; print(config.HOST)"') do set HOST=%%H
python -m uvicorn app.main:app --host %HOST% --port %PORT%
popd
echo Stopped. Run start_background.bat to bring the background app back.
pause
