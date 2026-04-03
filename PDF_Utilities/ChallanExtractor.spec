# -*- mode: python ; coding: utf-8 -*-

# Essential modules for fitz (PyMuPDF) and PIL to work. 
# We removed structural/binary modules from this list to prevent the "Missing Library" error.
minimal_excludes = [
    'unittest', 'doctest', 'pdb', 'profile', 'cProfile', 'pstats',
    'difflib', 'calendar', 'html', 'http', 'xmlrpc',
    'xml', 'mailbox', 'mimetypes', 'multiprocessing',
    'asyncio', 'concurrent', 'ctypes.test', 'lib2to3',
    'sqlite3', 'curses', 'readline', 'gettext',
    'ftplib', 'imaplib', 'smtplib', 'poplib', 'nntplib',
    'telnetlib', 'uuid', 'getpass', 'grp', 'nis', 'ossaudiodev',
    'spwd', 'sunau', 'termios', 'tty', 'pty', 'fcntl',
    'resource', 'syslog', 'aifc', 'audioop', 'chunk',
    'crypt', 'imghdr', 'sndhdr', 'xdrlib',
    # Heavy third-party libs
    'numpy', 'pandas', 'scipy', 'matplotlib', 'IPython',
    'jupyter', 'notebook', 'pytest', 'setuptools', 'pkg_resources',
    'pygments', 'docutils', 'sphinx', 'pydoc', 'PIL.ImageQt',
]

a = Analysis(
    ['challan_extractor.py'],
    pathex=[],
    binaries=[],
    datas=[],
    # Added 'pymupdf' to hiddenimports to ensure the backend is found
    hiddenimports=['PIL._tkinter_finder', 'fitz', 'pymupdf', 'PIL'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=minimal_excludes,
    noarchive=False,
    optimize=2,
)

pyz = PYZ(a.pure, optimize=2)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ChallanExtractor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,     # Removes symbols to save space
    upx=True,       # Set to True and ensure upx.exe is in your project folder
    upx_exclude=['vcruntime140.dll', 'msvcp140.dll', 'python3.dll'],
    runtime_tmpdir=None,
    console=False,  # No background terminal
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)