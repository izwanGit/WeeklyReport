# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, copy_metadata
import os

block_cipher = None
base_dir = os.path.abspath(os.getcwd())

datas = collect_data_files("streamlit")
datas += copy_metadata("streamlit")
datas += [("app.py", "."), ("template.html", "."), ("PETRONAS_LOGO_SQUARE.png", "."), ("PETRONAS_LOGO_HORIZONTAL.svg", ".")]

a = Analysis(
    ["run_app.py"],
    pathex=[base_dir],
    datas=datas,
    hiddenimports=["pandas", "openpyxl", "jinja2", "streamlit", "win32com.client", "pythoncom"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="PetronasReportTool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon=None
)
coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=True, upx_exclude=[],
    name="PetronasReportTool",
)
