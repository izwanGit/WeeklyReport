@echo off
setlocal enabledelayedexpansion

echo ==================================================
echo  PETRONAS Report Hub - Windows Build Kit
echo ==================================================
echo.

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+ and add it to PATH.
    pause
    exit /b 1
)

echo [1/5] Creating Virtual Environment...
if exist venv rmdir /s /q venv
python -m venv venv
if errorlevel 1 (
    echo [ERROR] Failed to create virtual environment.
    pause
    exit /b 1
)

call venv\Scripts\activate
if errorlevel 1 (
    echo [ERROR] Failed to activate virtual environment.
    pause
    exit /b 1
)

echo [2/5] Upgrading pip...
python -m pip install --upgrade pip --quiet

echo [3/5] Installing Requirements...
pip install --upgrade -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install requirements.
    pause
    exit /b 1
)

echo [4/5] Building EXE (this takes 3-8 minutes)...
pyinstaller --clean --noconfirm app.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed. Check the output above for details.
    pause
    exit /b 1
)

echo [5/5] Verifying build output...
if not exist "dist\PETRONAS Report Hub\PETRONAS Report Hub.exe" (
    echo [ERROR] Build output not found - something went wrong.
    pause
    exit /b 1
)

echo.
echo ==================================================
echo  BUILD COMPLETE!
echo  Your application is in:
echo    dist\PETRONAS Report Hub\
echo.
echo  Send the ENTIRE FOLDER (not just the .exe) to Windows users.
echo  Double-click 'PETRONAS Report Hub.exe' to launch.
echo ==================================================
echo.
pause
