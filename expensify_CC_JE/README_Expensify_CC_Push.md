# Expensify Credit Card Push Automation

Automates the preparation of Expensify credit card journal-entry push files for PPS, ESR, and Operating Husky.

This script takes a single Expensify transaction workbook, removes non-project expenses, validates required department and project IDs, creates three journal-entry import worksheets, and exports each journal entry as a separate CSV file for upload.

## Why This Script Exists

Preparing intercompany and project-expense journal entries manually can be repetitive and error-prone. The process may require:

- Reviewing every Expensify transaction
- Removing expenses that do not belong to the intended project account range
- Confirming that each transaction has a department and project ID
- Reversing debit and credit activity between subsidiaries
- Calculating balancing entries
- Formatting data into a NetSuite-compatible journal import layout
- Creating separate import files for each entity

This script standardizes that workflow and performs the repetitive processing automatically.

It is intended to:

- Reduce manual journal-entry preparation time
- Improve consistency across monthly reporting periods
- Identify missing department or project information before import
- Reduce debit, credit, and balancing-entry errors
- Preserve the original workbook by creating a processed copy
- Produce ready-to-review Excel and CSV output files

## What the Script Does

The script performs the following steps:

1. Prompts for the journal-entry date.
2. Prompts for the reporting period.
3. Finds the only `.xlsx` workbook located in the same folder as the script.
4. Copies the source workbook to a new output workbook.
5. Renames the first worksheet to `Transactions`.
6. Removes transaction rows whose account in Column D does not begin with `12`.
7. Reviews the remaining rows for missing department IDs in Column R.
8. Reviews the remaining rows for missing project IDs in Column J.
9. Prompts the user to enter missing internal IDs when needed.
10. Inserts a memo formula into a new Column S.
11. Creates the following journal-entry worksheets:
    - `PPS Push to ESR`
    - `ESR Push for Exp`
    - `OH Push for Exp`
12. Calculates the required balancing journal-entry lines.
13. Exports each push worksheet as a separate CSV file.
14. Saves the completed Excel workbook.

## Input Workbook Requirements

The folder containing the script must contain exactly one `.xlsx` workbook.

The script uses the first worksheet in that workbook and expects transaction data to begin on Row 2.

### Required Columns

| Column | Purpose |
|---|---|
| A | Used to determine whether a transaction row contains data |
| B | Fixed asset GL or identifying information displayed when a row is removed |
| C | Used in the transaction memo |
| D | Journal-entry line account; only rows beginning with `12` are retained |
| F | Transaction date used in the transaction memo |
| G | Used in the transaction memo |
| H | Used in the transaction memo |
| J | Project internal ID |
| K | Used in the transaction memo |
| L | Debit activity for the ESR journal and credit activity for the PPS journal |
| M | Credit activity for the ESR journal and debit activity for the PPS journal |
| R | Department internal ID |

The workbook should not contain blank rows within the transaction data. Processing stops when the script reaches the first blank cell in Column A.

## Generated Journal-Entry Columns

Each push worksheet contains the following headers:

| Column | Header |
|---|---|
| A | Entry No. |
| B | subsidiary |
| C | memo |
| D | trandate |
| E | journalItemLine_account |
| F | journalItemLine_debitAmount |
| G | journalItemLine_creditAmount |
| H | journalItemLine_memo |
| I | constituent_project |
| J | constituent_department |

## Journal Entries Created

### PPS Push to ESR

Creates transaction lines for subsidiary `69`.

For each retained transaction:

- Column M is used as the debit amount.
- Column L is used as the credit amount.
- Account `21000` is used for the balancing line.
- Department `202` is assigned to the balancing line.

### ESR Push for Exp

Creates transaction lines for subsidiary `11`.

For each retained transaction:

- Column L is used as the debit amount.
- Column M is used as the credit amount.
- Account `31100` is used for the balancing line.
- Department `202` is assigned to the balancing line.

### OH Push for Exp

Creates a two-line journal entry for subsidiary `10`.

The journal uses the difference between the total of Column L and the total of Column M.

- Account `21470` receives the credit.
- Account `16020` receives the debit.
- Department `202` is assigned to both lines.

## Requirements

- Windows, macOS, or Linux
- Python 3.10 or newer recommended
- Microsoft Excel or another application capable of opening `.xlsx` files
- The Python package `openpyxl`

Install the required package with:

```bash
pip install openpyxl
```

## Recommended Folder Setup

Place the Python script and the monthly Expensify workbook in the same folder.

Example:

```text
Expensify-CC-Push/
├── expensify_CC_Push_V16.py
└── Expensify_Transactions_January_2026.xlsx
```

There must be only one `.xlsx` file in the folder before the script is run.

Temporary Excel files beginning with `~$` are ignored.

## How to Run the Script

### 1. Install Python

Download and install Python from the official Python website.

During installation on Windows, select:

```text
Add Python to PATH
```

### 2. Install the Required Package

Open Command Prompt, PowerShell, or Terminal and run:

```bash
pip install openpyxl
```

### 3. Prepare the Folder

Place these items together in the same folder:

- `expensify_CC_Push_V16.py`
- One Expensify `.xlsx` transaction workbook

Close the Excel workbook before running the script.

### 4. Open a Terminal in the Folder

On Windows, open the folder in File Explorer, click the address bar, type `cmd`, and press Enter.

You can also navigate to the folder manually:

```bash
cd "C:\Path\To\Expensify-CC-Push"
```

### 5. Run the Script

```bash
python expensify_CC_Push_V16.py
```

If your computer uses the `py` launcher:

```bash
py expensify_CC_Push_V16.py
```

### 6. Enter the Requested Information

The script prompts for the journal-entry date:

```text
What is the JE Date (MM/DD/YYYY):
```

Example:

```text
01/31/2026
```

The script then prompts for the reporting period:

```text
What is the Reporting Period (Ex: January 2026)
```

Example:

```text
January 2026
```

### 7. Resolve Missing IDs

When a retained transaction is missing a department or project internal ID, the script displays identifying information from the row and asks for the correct numeric ID.

Only whole-number numeric IDs are accepted.

Example:

```text
Please provide the correct department internal ID for this expense:
```

The ID entered is written into the processed workbook.

## Output Files

The completed workbook is saved in the same folder using this naming format:

```text
PPS_OH_ESR_Expensify_CC_PUSH_<Reporting_Period>.xlsx
```

Example:

```text
PPS_OH_ESR_Expensify_CC_PUSH_January_2026.xlsx
```

The script also creates three CSV files:

```text
PPS Push to ESR_<Reporting_Period>.csv
ESR Push for Exp_<Reporting_Period>.csv
OH Push for Exp_<Reporting_Period>.csv
```

Example:

```text
PPS Push to ESR_January_2026.csv
ESR Push for Exp_January_2026.csv
OH Push for Exp_January_2026.csv
```

## Important Processing Rules

- The source workbook is copied before changes are made.
- Only one `.xlsx` source file may be present in the folder.
- Rows are retained only when the value in Column D begins with `12`.
- Processing stops at the first blank value in Column A.
- Missing project and department values require manual input.
- Project and department responses must be numeric internal IDs.
- Existing push worksheets with the same names are deleted and rebuilt.
- Reporting-period text is sanitized before being used in filenames.
- Amounts that are blank, invalid, or nonnumeric are treated as zero.
- Dates used in transaction memos are formatted as `MM/DD/YYYY` when recognized.

## Error Messages

### No Excel File Found

```text
No .xlsx file was found in the same folder as the Python script.
```

Place one source workbook in the same folder as the script.

### More Than One Excel File Found

```text
More than one .xlsx file was found in the same folder as the Python script.
Please keep only one .xlsx file in that folder.
```

Move prior output workbooks or unrelated Excel files to another folder before rerunning the script.

### Permission Denied

This commonly occurs when the source or output workbook is open in Excel.

Close the workbook and run the script again.

### Missing Worksheet or Unexpected Columns

The script uses the first worksheet and relies on fixed Excel column locations. Confirm that the source report follows the expected layout described in this README.

## Review Checklist

Before importing the generated CSV files:

- Confirm the journal-entry date.
- Confirm the reporting period in the journal memos.
- Review any manually entered project IDs.
- Review any manually entered department IDs.
- Confirm the expected transactions were retained.
- Confirm fixed-asset or non-project rows were removed.
- Confirm debit and credit totals balance.
- Confirm subsidiary, account, project, and department values.
- Retain the processed Excel workbook as supporting documentation.

## Repository Description

**Suggested GitHub repository description:**

> Automates Expensify credit card project-expense processing by validating transaction coding and generating balanced PPS, ESR, and Operating Husky journal-entry import files.

## Disclaimer

This tool is designed for a specific accounting workflow and workbook structure. Review all generated journal entries before uploading them into an accounting or ERP system.

The script does not replace accounting review, approval controls, or reconciliation procedures.
