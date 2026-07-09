@echo off
REM Drag a ZIP file onto this BAT file, or double-click it and paste the ZIP path.

cd /d "%~dp0"

if "%~1"=="" (
    python invoice_pdf_combiner.py
) else (
    python invoice_pdf_combiner.py "%~1"
)

pause
