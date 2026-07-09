@echo off
REM Builds a standalone EXE using PyInstaller.
REM Run this after installing requirements:
REM pip install -r requirements.txt

cd /d "%~dp0"

pyinstaller --onefile --name InvoicePDFCombiner invoice_pdf_combiner.py

echo.
echo Build complete.
echo Your EXE should be in the "dist" folder:
echo dist\InvoicePDFCombiner.exe
echo.
pause
