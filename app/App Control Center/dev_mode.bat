@echo off
rem Development: stops the background app, runs start_dashboard.bat (with --reload),
rem and starts the background app again when you close the dev server.
call "%~dp0stop_background.bat"
call "%~dp0..\..\start_dashboard.bat"
call "%~dp0start_background.bat"
