@echo off
title MyGenie Cookie Bridge + Report App

echo ============================================
echo   PETRONAS Weekly Report — Quick Launcher
echo ============================================
echo.
echo Starting Cookie Bridge receiver...
echo Starting Streamlit report app...
echo.
echo Both will open in this window.
echo Close this window to stop everything.
echo.

:: Start the cookie receiver in a separate window
start "Cookie Bridge" cmd /k "python "%~dp0cookie_bridge_server\cookie_receiver.py""

:: Wait 1 second then start Streamlit
timeout /t 1 /nobreak >nul

:: Start Streamlit — adjust the path to your app if needed
start "Report App" cmd /k "streamlit run "%~dp0app\weekly_report.py""

echo Both servers launched!
echo.
echo Next steps:
echo   1. Wait for the browser to open the Streamlit app
echo   2. Click the teal extension icon in your Edge toolbar
echo   3. Click "Send Session to Report App"
echo   4. Watch the counts auto-fill!
echo.
pause
