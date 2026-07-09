# Invoice PDF Combiner

A lightweight Windows automation tool for quickly processing invoice packets received by email.

This tool takes a ZIP file of invoice documents, extracts the contents, converts Excel spreadsheets to PDF, prioritizes invoice files first, and combines all supported documents into one clean PDF packet.

It is designed for Accounts Payable workflows where vendors or internal teams send multiple invoice-related files together, such as PDFs, backup schedules, approvals, and Excel workbooks. Instead of manually extracting, converting, sorting, and combining the documents, this tool creates a completed invoice packet in seconds.

## Repository Description

Windows automation tool that extracts invoice ZIP files, converts Excel backups to PDF, prioritizes invoice files first, and combines everything into one organized PDF packet.

## Why This Is Useful

Invoice processing often involves receiving multiple documents through email, downloading them, opening a ZIP file, reviewing the files, converting Excel backups to PDF, and manually combining everything into a single invoice packet.

This tool turns that process into a quick drag-and-drop workflow.

Instead of spending several minutes per invoice packet preparing files manually, the user can provide a ZIP file and have the final combined PDF generated automatically. This is especially helpful for high-volume AP workflows where small repetitive tasks add up quickly across dozens or hundreds of invoices.

## Efficiency Impact

This automation is a strong time saver because it reduces several manual steps into one repeatable process.

Manual invoice packet preparation often requires:

1. Downloading the ZIP file from email.
2. Extracting the ZIP folder.
3. Identifying the invoice document.
4. Opening Excel files.
5. Saving Excel files as PDFs.
6. Rearranging files so the invoice appears first.
7. Combining multiple documents into one PDF.
8. Renaming and saving the final packet.

This tool handles most of that automatically.

The result is faster invoice processing, fewer clicks, more consistent file organization, and less risk of missing backup documents during manual combination.

This automation takes a 2 minute task to 2 seconds.

## Best Use Case

This tool is ideal when an invoice is received by email with multiple supporting documents attached or zipped together.

Example ZIP contents:

- Vendor invoice PDF
- Excel backup schedule
- Supporting documentation
- Receipts
- Approval files
- Payment backup
- Project or cost detail files

The final output is one organized PDF that is easier to save, attach, review, upload, or route for approval.

## What It Does

- Accepts a ZIP file as the input.
- Automatically extracts the ZIP file.
- Finds all supported PDF and Excel files.
- Converts Excel spreadsheets into PDFs using Microsoft Excel.
- Combines all PDFs into one final document.
- Prioritizes files with `invoice` or `inv` in the filename so the invoice appears first.
- Saves the completed PDF to the user's Downloads folder.
- Creates logs for troubleshooting and auditability.

## Supported File Types

Currently supported:

- `.pdf`
- `.xlsx`
- `.xlsm`
- `.xls`

Excel files are converted to PDF before being added to the combined packet.

## Invoice Page Ordering

Files with invoice-like names are placed first in the combined PDF.

Examples that will be prioritized:

- `invoice_123.pdf`
- `Vendor Invoice.pdf`
- `vendor_invoice_123.pdf`
- `INV-123.xlsx`
- `123_inv.pdf`

Everything else is added afterward alphabetically.

## Output

The combined PDF is saved to the user's Windows Downloads folder.

Example output:

```text
C:\Users\YourName\Downloads\Vendor_Invoices_Combined.pdf
```

The output file is named after the original ZIP file.

Example input:

```text
Vendor Invoice Packet.zip
```

Example output:

```text
Vendor Invoice Packet_Combined.pdf
```

If a file with the same name already exists, the script creates a numbered version instead of overwriting it.

Example:

```text
Vendor Invoice Packet_Combined_2.pdf
```

## Requirements

This tool is intended for Windows users.

Required:

- Python
- Microsoft Excel installed
- Python packages listed in `requirements.txt`

Python packages:

```text
pypdf
pywin32
pyinstaller
```

## Getting Started

### 1. Install Python

Use your existing Python install if it is already working.

To check, open Command Prompt and run:

```bat
python --version
```

### 2. Download or clone the repository

Download the repository files or clone the repo to your computer.

Example location:

```text
C:\Users\YourName\Desktop\Python\PDF Combiner
```

### 3. Open Command Prompt in the project folder

In File Explorer, open the project folder.

Click the address bar, type:

```text
cmd
```

Then press Enter.

### 4. Install required packages

Open Command Prompt in this folder and run:

```bat
pip install -r requirements.txt
```

### 5. Test the script

Run:

```bat
python invoice_pdf_combiner.py
```

Then paste the path to a ZIP file when prompted.

Example:

```text
C:\Users\YourName\Downloads\Vendor Invoice Packet.zip
```

Or drag a ZIP file onto:

```text
Run_Invoice_PDF_Combiner.bat
```

### 6. Build the EXE

Run:

```bat
Build_EXE.bat
```

After it finishes, your EXE will be here:

```text
dist\InvoicePDFCombiner.exe
```

### 7. Use the EXE

Drag a ZIP file onto:

```text
InvoicePDFCombiner.exe
```

The combined PDF will save to your Windows Downloads folder.

## Daily Workflow

Once set up, the workflow is simple:

1. Receive or download an invoice ZIP file from email.
2. Run the script, BAT file, or EXE.
3. Paste or drag in the ZIP file path.
4. The tool extracts, converts, sorts, and combines the files.
5. Find the completed PDF in your Downloads folder.
6. Save, upload, attach, or route the combined invoice packet as needed.

## Notes

- This requires Microsoft Excel to be installed for Excel-to-PDF conversion.
- Excel files use their existing workbook print settings when exported to PDF.
- If an Excel file has messy print areas, the resulting PDF may also look messy.
- The script currently does not process Word documents, image files, or email files.
- Logs are saved in the `Logs` folder.
- The tool is designed for quick invoice packet processing, not full document management.

## Future Improvement Ideas

Possible future enhancements:

- Add support for Word documents.
- Add support for image files such as `.png` and `.jpg`.
- Add a simple graphical interface.
- Add a watched folder where ZIP files can be dropped and processed automatically.
- Add custom naming rules by vendor or invoice number.
- Add better invoice detection using OCR or PDF text search.
