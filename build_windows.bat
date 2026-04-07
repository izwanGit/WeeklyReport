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
pyinstaller --clean windows_build.spec

echo.
echo ==================================================
echo BUILD COMPLETE!
echo Your executable is in: dist\PetronasReportTool\
echo Double-click 'PetronasReportTool.exe' to run.
echo ==================================================
pause
