@echo off
rem Removes the Windows auto-start shortcut. Does not stop the app running now (use stop_background.bat).
set "LNK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Credit Department.lnk"
if exist "%LNK%" (
    del "%LNK%"
    echo Auto-start removed.
) else (
    echo Auto-start was not installed.
)
timeout /t 3 >nul
