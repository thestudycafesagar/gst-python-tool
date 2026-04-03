# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all

# ── Smart source collector — skips venvs, dist, build, output, caches ────────
_SKIP_DIRS = frozenset({
    '.venv', 'venv', 'env', '.env', 'Venv', 'ENV', 'virtualenv',
    'dist', 'build', 'output',
    '__pycache__', '.git', '.idea', '.vscode', '.claude',
    'GST_Downloads', 'GST_3B_Downloads', 'IMS_Downloads',
    'temp', 'tmp', 'logs', 'node_modules',
    'site-packages', 'Lib', 'Scripts', 'Include',  # venv internals
})

# Absolute paths to known venvs — belt-and-suspenders exclusion
_SKIP_ABS_PATHS = {
    os.path.abspath('Bank Statement To Excel/venv'),
    os.path.abspath('Bank Statement To Excel/.venv'),
    os.path.abspath('Email-Tools/.venv'),
    os.path.abspath('Email-Tools/venv'),
    os.path.abspath('PDF_Utilities/.venv'),
    os.path.abspath('PDF_Utilities/venv'),
}
_SKIP_EXTS = frozenset({
    '.exe', '.dll', '.pyd', '.so',
    '.pyc', '.pyo', '.pycache',
    '.db', '.sqlite', '.log',
})

def collect_tool_sources(tool_dir, dest_name):
    """Return (abs_src, dest_folder) pairs for source files only — no venv/dist/build."""
    pairs = []
    tool_dir = os.path.abspath(tool_dir)
    for dirpath, dirnames, filenames in os.walk(tool_dir):
        abs_dirpath = os.path.abspath(dirpath)
        # Skip known venv absolute paths
        if any(abs_dirpath.startswith(vp) for vp in _SKIP_ABS_PATHS):
            dirnames[:] = []
            continue
        # Skip by directory name
        dirnames[:] = [d for d in dirnames
                       if d not in _SKIP_DIRS
                       and not os.path.abspath(os.path.join(dirpath, d)) in _SKIP_ABS_PATHS]
        rel = os.path.relpath(dirpath, tool_dir)
        dest_sub = dest_name if rel == '.' else (dest_name + '/' + rel.replace('\\', '/'))
        for fname in filenames:
            if os.path.splitext(fname.lower())[1] in _SKIP_EXTS:
                continue
            pairs.append((os.path.join(dirpath, fname), dest_sub))
    return pairs

# ── Base datas (assets + updater) ────────────────────────────────────────────
datas = [
    ('studycafelogo.ico', '.'),
    ('studycafelogo.png', '.'),
    ('dist/StudyCafeSuite_Updater.exe', '.'),   # bundled updater for auto-update
]

# Add tool sources — only .py / .json / .xlsx / .png / etc., NO venvs or dist EXEs
for _tool_dir, _dest in [
    ('GST',                     'GST'),
    ('Income Tax',              'Income Tax'),
    ('PDF_Utilities',           'PDF_Utilities'),
    ('Bank Statement To Excel', 'Bank Statement To Excel'),
    ('Email-Tools',             'Email-Tools'),
    ('GST_RECO',                'GST_RECO'),
]:
    datas += collect_tool_sources(_tool_dir, _dest)

# ── Hidden imports & package collection ──────────────────────────────────────
binaries = []
hiddenimports = [
    'customtkinter', 'darkdetect',
    'PIL', 'PIL.Image', 'PIL.ImageTk',
    'fitz', 'pdfplumber', 'pdfminer',
    'selenium', 'webdriver_manager',
    'pandas', 'numpy', 'openpyxl',
    'win32com', 'win32com.client', 'pythoncom', 'pywintypes', 'win32api', 'win32con', 'win32gui',
    # tkinter submodules not auto-detected by PyInstaller
    'tkinter.scrolledtext',   # used by PDF_Utilities and Email-Tools
    'tkinter.colorchooser',   # used by PDF_Utilities (Redact tool)
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.simpledialog',
    'tkinter.font',
]
for pkg in ('customtkinter', 'darkdetect', 'selenium', 'webdriver_manager',
            'pdfplumber', 'pdfminer', 'fitz'):
    tmp = collect_all(pkg)
    datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ['GST_Suite.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Heavy data-science / dev tools — not used by GST Suite
        'matplotlib', 'matplotlib.pyplot', 'mpl_toolkits',
        'scipy', 'sklearn', 'scikit_learn',
        'IPython', 'ipython', 'ipykernel', 'ipywidgets',
        'notebook', 'nbformat', 'nbconvert', 'jupyterlab', 'jupyter_client',
        'pytest', '_pytest',
        # Web frameworks — not needed at runtime
        'streamlit',          # used only in Bank Statement venv, NOT in main app
        'flask', 'werkzeug',
        'django',
        'fastapi', 'uvicorn', 'starlette',
        # Computer vision / OCR — only in Bank Statement venv
        'cv2',
        'pytesseract',
        'pdf2image',
        'skimage',
        # Cloud / infra SDKs
        'boto3', 'botocore', 's3transfer',
        'google.cloud',
        # Database drivers not used in the bundled app
        'psycopg2', 'pymysql', 'pyodbc', 'cx_Oracle',
        # Documentation / packaging tools
        'docutils', 'sphinx',
        'setuptools', 'pkg_resources._vendor',
        # Unused stdlib
        'turtle', 'tkinter.test', 'tkinter.tix', 'tkinter.dnd',
        'test', 'xmlrpc',
        # Unused image extras
        'imageio',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# ── Single-file EXE (windowed GUI; no console window) ───────────────────────
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='GST_Suite',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['studycafelogo.ico'],
)
# No COLLECT — single-file EXE
