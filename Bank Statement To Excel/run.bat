@echo off
echo Starting Bank Statement to Excel Converter...
echo.

:: Activate venv if it exists
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

:: Add Poppler to PATH if installed at default location
if exist "C:\poppler\Library\bin" (
    set PATH=%PATH%;C:\poppler\Library\bin
)

streamlit run app.py --server.maxUploadSize 200

pause
