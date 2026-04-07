# 📊 Petronas Weekly Report Tool - Windows Build Instructions

This folder contains a **Build Kit** for turning your Streamlit dashboard into a standalone Windows `.exe`.

### 🚀 To Build the EXE (Run this on Windows):
1. Copy this entire project folder to your Windows machine.
2. Double-click the **`build_windows.bat`** file.
3. This will automatically create a virtual environment, install dependencies, and run PyInstaller.

### 📁 Your Executable Location:
Once the build is complete, you will find a new folder at:
`dist\PetronasReportTool\`
- Launch the **`PetronasReportTool.exe`** inside that folder.
- It will automatically start the server and open your default browser.

### 🛠️ Troubleshooting (If the build fails):
- Ensure you have **Python 3.9+** installed on your Windows machine.
- Make sure to keep the `app.py`, `run_app.py`, and `template.html` files in the same folder as the `.bat`.
