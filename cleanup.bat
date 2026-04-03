@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title GST Suite Cleanup
cd /d "%~dp0"

echo.
echo  ================================================
echo   GST Suite - Cleanup Utility
echo  ================================================
echo  This will delete:
echo    - build               (PyInstaller cache)
echo    - All __pycache__ folders
echo    - All *.pyc files
echo    - All *.log files
echo    - Old EXEs in dist (keeps GST_Suite.exe and Updater.exe)
echo    - Download output folders
echo.
echo  Your source .py files are NOT touched.
echo.
set /p CONFIRM=Type YES to proceed: 
if /i not "!CONFIRM!"=="YES" (
    echo  Cancelled.
    pause
    exit /b 0
)
echo.
echo  [1/6] Deleting build ...
if exist "build" (
    rmdir /s /q "build"
    echo         Deleted build
) else (
    echo         build not found - skipping.
)

echo  [2/6] Deleting __pycache__ folders...
for /d /r . %%d in (__pycache__) do (
    if exist "%%d" rmdir /s /q "%%d" 2>nul
)
echo         Done.

echo  [3/6] Deleting *.pyc files...
for /r . %%f in (*.pyc) do del /q "%%f" 2>nul
echo         Done.

echo  [4/6] Deleting *.log files...
for /r . %%f in (*.log) do del /q "%%f" 2>nul
echo         Done.

echo  [5/6] Cleaning old EXEs in dist ...
if exist "dist" (
    for /f "delims=" %%f in ('dir /b "dist\*.exe" 2^>nul') do (
        if /i not "%%f"=="GST_Suite.exe" (
            if /i not "%%f"=="StudyCafeSuite_Updater.exe" (
                del /q "dist\%%f"
                echo         Deleted dist\%%f
            )
        )
    )
    echo         Kept: dist\GST_Suite.exe and dist\StudyCafeSuite_Updater.exe
) else (
    echo         dist not found - skipping.
)

echo  [6/6] Deleting download output folders...
for %%d in (GST_Downloads GST_3B_Downloads IMS_Downloads output) do (
    if exist "%%d" (
        rmdir /s /q "%%d"
        echo         Deleted %%d
    )
)
echo         Done.

echo.
echo  ================================================
echo   Cleanup complete! Run build.bat for a fresh EXE.
echo  ================================================
echo.
pause
