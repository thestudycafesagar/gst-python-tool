# -*- mode: python ; coding: utf-8 -*-
"""
updater.spec — PyInstaller spec for the GST Suite updater.

Build:  pyinstaller updater.spec --noconfirm --clean
Output: dist/updater.exe  (console EXE, ~8 MB)

Build this BEFORE building GST_Suite.exe because GST_Suite.spec
bundles dist/updater.exe inside the main EXE.
"""

import os

SPEC_DIR = os.path.dirname(os.path.abspath(SPEC))   # noqa: F821

a = Analysis(
    [os.path.join(SPEC_DIR, "updater.py")],
    pathex=[SPEC_DIR],
    binaries=[],
    datas=[],
    hiddenimports=[
        "urllib.request",
        "shutil",
        "tempfile",
        "subprocess",
        "argparse",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter", "customtkinter", "pandas", "numpy",
        "selenium", "PIL", "fitz", "win32com",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)   # noqa: F821

exe = EXE(          # noqa: F821
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="StudyCafeSuite_Updater",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,               # show progress in a CMD window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
