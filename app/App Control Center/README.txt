APP CONTROL CENTER - how to run Credit Department

install_autostart.bat    Run once. The app starts hidden at every Windows login (tray icon, auto-restart on crash).
uninstall_autostart.bat  Removes the auto-start.
start_background.bat     Start the hidden background app now (no command window).
stop_background.bat      Stop the hidden background app.
run_terminal.bat         Visible terminal, no reload (stops the background app first).
dev_mode.bat             Edit code with --reload (stops background app, restores it afterwards).
background_tray.pyw      The tray app itself (.pyw = runs with no console window).

The address and port come from config.py (HOST / PORT). Only one mode can use the port at a time.
After code edits the background app needs "Restart server" from its tray menu (right-click the CD icon).
Server log: logs\server.log.  start_dashboard.bat in the project folder is unchanged (visible window).
