# -*- mode: python ; coding: utf-8 -*-
# ──────────────────────────────────────────────────────────────
# PETRONAS Report Hub — PyInstaller spec
# Built for Windows (one-folder / COLLECT build)
# ──────────────────────────────────────────────────────────────

from PyInstaller.utils.hooks import collect_data_files, copy_metadata, collect_submodules

# ── Data files ─────────────────────────────────────────────────
datas = []

# Project assets
datas += [
    ('template.html',                    '.'),
    ('Report_Hub.py',                    '.'),
    ('pages',                            'pages'),
    ('icon_transparent.ico',             '.'),
    ('PETRONAS_LOGO_SQUARE.png',         '.'),
    ('PETRONAS_LOGO_HORIZONTAL.svg',     '.'),
    ('PETRONAS_LOGO_HORIZONTAL_WHITE.svg', '.'),
    ('.streamlit',                       '.streamlit'),
]

# If template.pptx exists, bundle it too
import os
if os.path.exists('template.pptx'):
    datas += [('template.pptx', '.')]

# Streamlit — must bundle its entire static web UI + data files
datas += collect_data_files('streamlit')

# pptx (python-pptx) uses templates/themes stored in package files
datas += collect_data_files('pptx')

# altair/vega_datasets used by Streamlit for some charts
try:
    datas += collect_data_files('altair')
except Exception:
    pass

try:
    datas += collect_data_files('vega_datasets')
except Exception:
    pass

# jinja2 uses template files
datas += collect_data_files('jinja2')

# ── Package metadata (importlib.metadata lookups by Streamlit) ─
datas += copy_metadata('streamlit')
datas += copy_metadata('pandas')
datas += copy_metadata('numpy')
datas += copy_metadata('openpyxl')
datas += copy_metadata('python-pptx')
datas += copy_metadata('pymupdf')
datas += copy_metadata('jinja2')
datas += copy_metadata('altair')
datas += copy_metadata('click')
datas += copy_metadata('packaging')
datas += copy_metadata('pyarrow')
datas += copy_metadata('PIL')   # Pillow

# ── Hidden imports ─────────────────────────────────────────────
hiddenimports = [
    # Streamlit internals
    'streamlit',
    'streamlit.web.cli',
    'streamlit.web.server',
    'streamlit.runtime',
    'streamlit.runtime.scriptrunner',

    # Data
    'pandas',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.offsets',
    'numpy',
    'pyarrow',
    'openpyxl',
    'openpyxl.styles.stylesheet',

    # Jinja2
    'jinja2',
    'jinja2.ext',

    # PPTX / PDF
    'pptx',
    'pptx.util',
    'pptx.oxml',
    'pptx.oxml.ns',
    'pymupdf',
    'fitz',

    # Windows COM (optional — no error if missing on non-Windows build)
    'win32com',
    'win32com.client',
    'pythoncom',
    'pywintypes',

    # Streamlit optional image / chart libs
    'PIL',
    'PIL.Image',
    'altair',
    'click',
    'packaging',
    'packaging.version',
    'packaging.specifiers',
    'packaging.requirements',
    'validators',
    'tornado',
    'tornado.web',
    'tornado.ioloop',
    'tornado.websocket',
]

hiddenimports += collect_submodules('streamlit')
hiddenimports += collect_submodules('altair')

block_cipher = None

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',   # large and unused
        'scipy',
        'sklearn',
        'tensorflow',
        'torch',
        'cv2',
        'PyQt5',
        'tkinter',
    ],
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
    # Keep console=True so startup errors appear in the window
    # Set to False once you have confirmed the app runs cleanly on Windows
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
