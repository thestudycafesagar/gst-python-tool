# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all


_SKIP_DIRS = frozenset({
    '.venv', 'venv', 'env', '.env', 'Venv', 'ENV', 'virtualenv',
    'dist', 'build', 'output',
    '__pycache__', '.git', '.idea', '.vscode', '.claude',
    'GST_Downloads', 'GST_3B_Downloads', 'IMS_Downloads',
    'temp', 'tmp', 'logs', 'node_modules',
    'site-packages', 'Lib', 'Scripts', 'Include',
})

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
    pairs = []
    tool_dir = os.path.abspath(tool_dir)

    for dirpath, dirnames, filenames in os.walk(tool_dir):
        abs_dirpath = os.path.abspath(dirpath)

        # Skip known virtual environments completely.
        if any(abs_dirpath.startswith(vp) for vp in _SKIP_ABS_PATHS):
            dirnames[:] = []
            continue

        # Skip unwanted folders.
        dirnames[:] = [
            d for d in dirnames
            if d not in _SKIP_DIRS
            and os.path.abspath(os.path.join(dirpath, d)) not in _SKIP_ABS_PATHS
        ]

        rel = os.path.relpath(dirpath, tool_dir)
        dest_sub = dest_name if rel == '.' else (dest_name + '/' + rel.replace('\\', '/'))

        for fname in filenames:
            # Skip Office temp lock files, which can be locked during build.
            if fname.startswith('~$'):
                continue
            if os.path.splitext(fname.lower())[1] in _SKIP_EXTS:
                continue
            pairs.append((os.path.join(dirpath, fname), dest_sub))

    return pairs


import customtkinter as _ctk
_ctk_dir = os.path.dirname(_ctk.__file__)

datas = [
    ('studycafelogo.ico', '.'),
    ('studycafelogo.png', '.'),
    ('dist/StudyCafeSuite_Updater.exe', '.'),
    (_ctk_dir, 'customtkinter'),
]

for _tool_dir, _dest in [
    ('GST', 'GST'),
    ('Income Tax', 'Income Tax'),
    ('PDF_Utilities', 'PDF_Utilities'),
    ('Bank Statement To Excel', 'Bank Statement To Excel'),
    ('Outlook Email Tools', 'Outlook Email Tools'),
    ('Gmail-Tools', 'Gmail-Tools'),
    ('tally tool', 'tally tool'),
    ('GST_RECO', 'GST_RECO'),
]:
    datas += collect_tool_sources(_tool_dir, _dest)


binaries = []
hiddenimports = [
    'customtkinter', 'darkdetect',
    'PIL', 'PIL.Image', 'PIL.ImageTk',
    'fitz', 'pdfplumber', 'pdfminer',
    'selenium', 'webdriver_manager',
    'pandas', 'numpy', 'openpyxl',
    'requests', 'urllib3', 'charset_normalizer', 'certifi', 'idna',
    'win32com', 'win32com.client', 'pythoncom', 'pywintypes', 'win32api', 'win32con', 'win32gui',
    'smtplib',
    'email', 'email.message', 'email.mime', 'email.mime.multipart',
    'email.mime.text', 'email.mime.base', 'email.mime.application',
    'email.encoders', 'email.utils', 'email.header',
    'xml', 'xml.etree', 'xml.etree.ElementTree', 'xml.dom', 'xml.dom.minidom',
    'tkinter.scrolledtext',
    'tkinter.colorchooser',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.simpledialog',
    'tkinter.font',
    'stealth_driver',
    'pypdf',
]

for pkg in ('customtkinter', 'darkdetect', 'selenium', 'webdriver_manager',
            'pdfplumber', 'pdfminer', 'fitz'):
    tmp = collect_all(pkg)
    datas += tmp[0]
    binaries += tmp[1]
    hiddenimports += tmp[2]


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
        'matplotlib', 'matplotlib.pyplot', 'mpl_toolkits',
        'scipy', 'sklearn', 'scikit_learn',
        'IPython', 'ipython', 'ipykernel', 'ipywidgets',
        'notebook', 'nbformat', 'nbconvert', 'jupyterlab', 'jupyter_client',
        'pytest', '_pytest',
        'streamlit',
        'flask', 'werkzeug',
        'django',
        'fastapi', 'uvicorn', 'starlette',
        'cv2',
        'pytesseract',
        'pdf2image',
        'skimage',
        'boto3', 'botocore', 's3transfer',
        'google.cloud',
        'psycopg2', 'pymysql', 'pyodbc', 'cx_Oracle',
        'docutils', 'sphinx',
        'setuptools', 'pkg_resources._vendor',
        'turtle', 'tkinter.test', 'tkinter.tix', 'tkinter.dnd',
        'test', 'xmlrpc',
        'imageio',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='AutomationCafe',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
