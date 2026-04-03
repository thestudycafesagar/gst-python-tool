@echo off
echo ============================================================
echo  GST Suite Updater — Build Script
echo  Builds StudyCafeSuite_Updater.exe which is bundled inside StudyCafeSuite.exe
echo ============================================================
echo.

echo [1/2] Installing dependencies...
pip install pyinstaller --upgrade --quiet
pip install requests --quiet

echo.
echo [2/2] Building StudyCafeSuite_Updater.exe...
pyinstaller updater.spec --noconfirm --clean

echo.
if exist "dist\StudyCafeSuite_Updater.exe" (
    echo  SUCCESS: dist\StudyCafeSuite_Updater.exe created.
    echo  Next step: run build.bat to build the main StudyCafeSuite.exe
) else (
    echo  ERROR: Build failed. Check the output above.
)
echo.
pause
