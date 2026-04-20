@echo off
echo ==================================================
echo Petronas Weekly Report Tool - Windows Build Kit
echo ==================================================
echo.
echo 1. Creating Virtual Environment...
python -m venv venv
call venv\Scripts\activate

echo 2. Installing Requirements...
pip install -r requirements.txt
pip install pyinstaller streamlit

echo 3. Generating EXE (This may take 1-2 minutes)...
pyinstaller --clean app.spec

echo.
echo ==================================================
echo BUILD COMPLETE!
echo Your executable is in: dist\PETRONAS Report Hub\
echo Double-click 'PETRONAS Report Hub.exe' to run.
echo ==================================================
pause
