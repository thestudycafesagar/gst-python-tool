@echo off
echo Building PDFTools.exe ...
cd /d "%~dp0"
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name PDFTools ^
    --clean ^
    --hidden-import pdfplumber ^
    --hidden-import pdfminer ^
    --hidden-import pdfminer.high_level ^
    --hidden-import pdfminer.layout ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.styles ^
    --hidden-import openpyxl.utils ^
    --hidden-import pikepdf ^
    --hidden-import PIL ^
    --hidden-import PIL.Image ^
    --collect-all pdfplumber ^
    --collect-all pdfminer ^
    --collect-all pikepdf ^
    --collect-all PIL ^
    main.py
echo.
if exist dist\PDFTools.exe (
    echo SUCCESS: dist\PDFTools.exe is ready!
    explorer dist
) else (
    echo FAILED: check the output above for errors.
)
pause
