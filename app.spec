# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import copy_metadata

datas = [
    ('template.html', '.'),
    ('Report_Hub.py', '.'),
    ('pages', 'pages'),
    ('icon_transparent.ico', '.'),
    ('PETRONAS_LOGO_SQUARE.png', '.'),
    ('PETRONAS_LOGO_HORIZONTAL.svg', '.'),
    ('PETRONAS_LOGO_HORIZONTAL_WHITE.svg', '.'),
    ('.streamlit/*', '.streamlit'),
]
datas += collect_data_files('streamlit')
datas += copy_metadata('streamlit')

block_cipher = None

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['streamlit', 'pandas', 'openpyxl', 'jinja2', 'win32com.client', 'pymupdf', 'pptx'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PETRONAS Report Hub',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon_transparent.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PETRONAS Report Hub',
)
