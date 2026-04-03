@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title StudyCafe Suite Builder
cd /d "%~dp0"

echo.
echo  ================================================
echo   StudyCafe Suite -- EXE Builder
echo  ================================================
echo.
echo  Working directory: %CD%
echo.

echo  [0/5] Checking for active virtual environment...
if defined VIRTUAL_ENV (
    echo  WARNING: A virtual environment is active: %VIRTUAL_ENV%
    echo  This can cause PyInstaller to bundle venv packages and create a 3GB EXE.
    echo  Please deactivate it first by running:  deactivate
    echo.
    set /p VENV_CONFIRM="Type YES to build anyway (not recommended): "
    if /i not "!VENV_CONFIRM!"=="YES" (
        echo  Aborted. Run 'deactivate' then re-run build.bat.
        pause
        exit /b 1
    )
)
echo.

echo  [1/5] Checking PyInstaller...
pip install --upgrade pyinstaller >nul 2>&1
if errorlevel 1 (
    echo  ERROR: pip failed. Make sure Python is on PATH.
    pause
    exit /b 1
)
echo         OK
echo.

echo  [2/5] Installing dependencies...
pip install --quiet customtkinter Pillow numpy pandas openpyxl selenium webdriver-manager pdfplumber pdfminer.six PyMuPDF pywin32 reportlab pikepdf pypdf pyinstaller
echo         OK
echo.

echo  [3/5] Checking updater...
if exist "dist\StudyCafeSuite_Updater.exe" (
    echo         StudyCafeSuite_Updater.exe already exists - skipping.
    echo         Delete dist\StudyCafeSuite_Updater.exe to force a rebuild.
) else (
    echo         Building StudyCafeSuite_Updater.exe...
    rmdir /s /q "build\updater" 2>nul
    pyinstaller updater.spec --noconfirm
    if errorlevel 1 (
        echo  ERROR: Updater build failed.
        pause
        exit /b 1
    )
    echo         Updater built OK.
)
echo.

echo  [4/5] Building GST_Suite.exe...
echo         This may take 2-5 minutes...
pyinstaller GST_Suite.spec --noconfirm --clean

if errorlevel 1 (
    echo.
    echo  BUILD FAILED - scroll up to see the error.
    pause
    exit /b 1
)
echo.

echo  [5/7] Verifying output...
if exist "dist\GST_Suite.exe" (
    echo         dist\GST_Suite.exe is ready - all tools are bundled inside.
) else (
    echo  ERROR: dist\GST_Suite.exe not found.
    pause
    exit /b 1
)
echo.

echo  [6/7] Syncing launcher copy...
copy /y "dist\GST_Suite.exe" "GST_Suite.exe" >nul
if errorlevel 1 (
    echo  WARNING: Could not update root GST_Suite.exe copy.
) else (
    echo         Root GST_Suite.exe updated from latest dist build.
)
if exist "dist\GST_Suite\GST_Suite.exe" (
    echo         NOTE: Old one-dir build detected at dist\GST_Suite\GST_Suite.exe
    echo               Launch dist\GST_Suite.exe or root GST_Suite.exe, not the old one-dir EXE.
)
echo.

echo  [7/7] Validating bundled Python sources...
set "_ARCHIVE_LIST=%TEMP%\gst_suite_archive_list.txt"
powershell -NoProfile -Command "pyi-archive_viewer 'dist\\GST_Suite.exe' -l | Out-File -FilePath '%_ARCHIVE_LIST%' -Encoding utf8" >nul 2>&1
if errorlevel 1 (
    echo  WARNING: Could not inspect EXE archive with pyi-archive_viewer.
) else (
    powershell -NoProfile -Command "if (Select-String -Path '%_ARCHIVE_LIST%' -SimpleMatch 'GST\\GST 2B Downloader\\main.py') { exit 0 } else { exit 1 }" >nul 2>&1
    if errorlevel 1 (
        echo  WARNING: Missing GST\GST 2B Downloader\main.py in bundle.
    ) else (
        echo         OK - GST\GST 2B Downloader\main.py
    )

    powershell -NoProfile -Command "if (Select-String -Path '%_ARCHIVE_LIST%' -SimpleMatch 'PDF_Utilities\\main.py') { exit 0 } else { exit 1 }" >nul 2>&1
    if errorlevel 1 (
        echo  WARNING: Missing PDF_Utilities\main.py in bundle.
    ) else (
        echo         OK - PDF_Utilities\main.py
    )

    powershell -NoProfile -Command "if (Select-String -Path '%_ARCHIVE_LIST%' -SimpleMatch 'Bank Statement To Excel\\bank_to_excel.py') { exit 0 } else { exit 1 }" >nul 2>&1
    if errorlevel 1 (
        echo  WARNING: Missing Bank Statement To Excel\bank_to_excel.py in bundle.
    ) else (
        echo         OK - Bank Statement To Excel\bank_to_excel.py
    )

    powershell -NoProfile -Command "if (Select-String -Path '%_ARCHIVE_LIST%' -SimpleMatch 'Email-Tools\\main.py') { exit 0 } else { exit 1 }" >nul 2>&1
    if errorlevel 1 (
        echo  WARNING: Missing Email-Tools\main.py in bundle.
    ) else (
        echo         OK - Email-Tools\main.py
    )
)
del "%_ARCHIVE_LIST%" >nul 2>&1
echo.

echo  ================================================
echo   BUILD COMPLETE!
echo   EXE is at: dist\GST_Suite.exe
echo   Root copy refreshed: GST_Suite.exe
echo   IMPORTANT: Use dist\GST_Suite.exe or root GST_Suite.exe
echo   Do NOT run dist\GST_Suite\GST_Suite.exe - stale one-dir build
echo  ================================================
echo.
pause
