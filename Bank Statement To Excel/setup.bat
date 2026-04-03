@echo off
echo ============================================
echo   Bank Statement to Excel - Setup Script
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.8+ from python.org
    pause
    exit /b 1
)

echo [1/3] Creating virtual environment...
python -m venv venv
if errorlevel 1 ( echo ERROR: Failed to create venv & pause & exit /b 1 )

echo [2/3] Activating virtual environment...
call venv\Scripts\activate.bat

echo [3/3] Installing dependencies...
pip install --upgrade pip
pip install streamlit pandas pdfplumber openpyxl Pillow pdf2image pytesseract numpy opencv-python-headless xlsxwriter

echo.
echo ============================================
echo   Setup complete!
echo ============================================
echo.
echo IMPORTANT - Additional requirements:
echo.
echo   1. Install Tesseract OCR (for scanned PDFs):
echo      Download from: https://github.com/UB-Mannheim/tesseract/wiki
echo      Default install path: C:\Program Files\Tesseract-OCR\tesseract.exe
echo.
echo   2. Install Poppler (for pdf2image):
echo      Download from: https://github.com/oschwartz10612/poppler-windows/releases
echo      Extract to C:\poppler and add C:\poppler\Library\bin to PATH
echo.
echo   To run the app:
echo      venv\Scripts\activate.bat
echo      streamlit run app.py
echo.
pause
