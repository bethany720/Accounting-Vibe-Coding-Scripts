# Invoice PDF Combiner

This tool takes a ZIP file, extracts it, converts Excel files to PDF, and combines all PDFs into one final PDF.

## Best daily workflow

Once built into an EXE:

1. Put `InvoicePDFCombiner.exe` somewhere easy, like your Desktop.
2. Drag a ZIP file onto it.
3. The final PDF will be saved in your Windows `Downloads` folder.

Example output:

`C:\Users\YourName\Downloads\Vendor_Invoices_Combined.pdf`

## What gets combined?

Supported files:

- `.pdf`
- `.xlsx`
- `.xlsm`
- `.xls`

Excel files are converted to PDF first using Microsoft Excel in the background.

## Invoice page ordering

Files with names containing invoice-like terms are placed first.

Examples that will be prioritized:

- `invoice_123.pdf`
- `Vendor Invoice.pdf`
- `INV-123.xlsx`
- `123_inv.pdf`

Everything else is added afterward alphabetically.

## Setup instructions

### 1. Install Python

Use your existing Python install if it is already working.

To check, open Command Prompt and run:

```bat
python --version
```

### 2. Install required packages

Open Command Prompt in this folder and run:

```bat
pip install -r requirements.txt
```

### 3. Test the script

Run:

```bat
python invoice_pdf_combiner.py
```

Then paste the path to a ZIP file when prompted.

Or drag a ZIP file onto:

```text
Run_Invoice_PDF_Combiner.bat
```

### 4. Build the EXE

Run:

```bat
Build_EXE.bat
```

After it finishes, your EXE will be here:

```text
dist\InvoicePDFCombiner.exe
```

### 5. Use the EXE

Drag a ZIP file onto:

```text
InvoicePDFCombiner.exe
```

The combined PDF will save to your Windows Downloads folder.

## Notes

- This requires Microsoft Excel to be installed for Excel-to-PDF conversion.
- If Excel files have messy print settings, the PDF may look messy too. The tool exports them using the workbook's current print setup.
- Logs are saved in the `Logs` folder.
- The script does not currently combine Word documents or image files. It can be expanded later.
