@echo off
rem Starts the hidden background app (tray icon) now, without waiting for the next login.
pushd "%~dp0..\.."
for /f "delims=" %%P in ('python -c "import sys, os; print(os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'))"') do set "PYW=%%P"
start "" "%PYW%" "%~dp0background_tray.pyw"
popd
echo Background app started. Look for the CD icon near the clock (it may be under the ^ arrow).
timeout /t 3 >nul
