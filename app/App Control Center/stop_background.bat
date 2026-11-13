@echo off
rem Stops THIS project's background tray app and its hidden server. The WalMart dashboard and visible terminal windows are not touched.
pushd "%~dp0..\.."
for /f "delims=" %%P in ('python -c "import config; print(config.PORT)"') do set "APP_PORT=%%P"
popd
powershell -NoProfile -ExecutionPolicy Bypass -Command "$port=$env:APP_PORT; Get-CimInstance Win32_Process | Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -like '*Credit Department*background_tray.pyw*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }; Get-CimInstance Win32_Process | Where-Object { $_.ProcessId -ne $PID -and $_.Name -eq 'pythonw.exe' -and $_.CommandLine -like '*uvicorn*app.main:app*' -and $_.CommandLine -like ('*--port ' + $port + '*') } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
echo Background app stopped.
timeout /t 2 >nul
