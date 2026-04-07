# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, copy_metadata, collect_submodules
import os

block_cipher = None
base_dir = os.path.abspath(os.getcwd())

datas = collect_data_files("streamlit")
datas += copy_metadata("streamlit")
datas += [("app.py", "."), ("template.html", "."), ("PETRONAS_LOGO_SQUARE.png", "."), ("PETRONAS_LOGO_HORIZONTAL.svg", "."), ("icon.ico", ".")]

all_hidden = ["pandas", "openpyxl", "jinja2", "win32com.client", "pythoncom"]
all_hidden += collect_submodules("streamlit")

a = Analysis(
    ["run_app.py"],
    pathex=[base_dir],
    datas=datas,
    hiddenimports=all_hidden,
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
    icon="icon.ico"
)
coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=True, upx_exclude=[],
    name="PetronasReportTool",
)
