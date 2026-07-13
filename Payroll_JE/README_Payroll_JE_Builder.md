# Payroll Journal Entry Builder

A Python automation tool that converts a payroll general ledger export into a formatted payroll journal-entry workbook and a NetSuite-ready CSV import file.

The script uses an account lookup table to group payroll activity, apply debit or credit signs, generate standardized memo descriptions, split selected accounts by pay component, and compare the completed journal entry against the payroll bank debit.

## Why This Tool Exists

Preparing a payroll journal entry manually can require repetitive account mapping, sign corrections, memo creation, pay-component analysis, spreadsheet formulas, CSV formatting, and reconciliation.

This tool automates those steps to:

- Reduce manual payroll journal-entry preparation time
- Apply account mappings consistently from one payroll period to the next
- Separate specified payroll accounts by pay component
- Flag unmapped accounts for review
- Produce a formatted Excel workpaper with supporting formulas
- Produce a CSV file in the expected journal-entry import structure
- Compare the journal-entry total with the bank debit before upload
- Retain the original payroll GL detail inside the completed workpaper

## What the Script Does

When run, the script:

1. Locates `Lookup.xlsx` in the same folder as the Python script.
2. Locates the single payroll GL workbook in that folder.
3. Prompts for the bank date and employee paid date.
4. Reads columns A through C from the payroll GL export.
5. Identifies the bank debit using the configured `10012.01:JP Morgan` line.
6. Stops processing GL detail at the bank debit line.
7. Excludes the configured `20600.07:Payroll Suspense` line.
8. Maps payroll accounts using `Lookup.xlsx`.
9. Splits accounts by pay component when a Pay Component value is provided in the lookup.
10. Creates a new `Payroll JE` worksheet in the original GL workbook.
11. Adds Excel formulas to calculate each journal-entry amount from the source GL detail.
12. Independently recalculates the journal-entry amounts in Python for the CSV export.
13. Creates a journal-entry import CSV.
14. Compares the journal-entry total with the bank debit.
15. Prints warnings for accounts or pay components that are not fully mapped.
16. Renames the original GL workbook using the payroll dates.

## Repository Files

```text
Payroll-JE-Builder/
├── build_payroll_journal_final_v29.py
├── Lookup_template.xlsx
└── README.md
```

Before running the script, make a working copy of `Lookup_template.xlsx` and rename it:

```text
Lookup.xlsx
```

The script specifically searches for a file named `Lookup.xlsx`.

## Requirements

- Windows, macOS, or Linux
- Python 3.10 or newer
- Microsoft Excel or another application capable of opening `.xlsx` files
- A payroll GL export in `.xlsx`, `.xlsm`, or `.xls` format

### Required Python Packages

```bash
pip install pandas numpy openpyxl
```

> `numpy` is currently imported by the script even though the primary processing is handled through pandas and openpyxl.

## Lookup File Structure

The first worksheet in `Lookup.xlsx` must contain these four columns in this order:

| Column | Header | Purpose |
|---|---|---|
| A | Account | Payroll GL account number to map |
| B | Memo | Description used in the journal-entry line memo |
| C | Sign | Amount multiplier, normally `1` or `-1` |
| D | Pay Component | Optional exact match used to split an account into separate journal-entry lines |

### Standard Account Mapping

For accounts that do not need to be split by pay component, leave **Pay Component** blank.

Example:

| Account | Memo | Sign | Pay Component |
|---:|---|---:|---|
| 1000.01 | EE and ER Taxes | -1 | |
| 1000.08 | Payroll | 1 | |

The script groups all source rows for that account into one journal-entry line.

### Pay-Component Mapping

Enter a value in **Pay Component** when an account must be divided based on the value in column B of the source GL workbook.

Example:

| Account | Memo | Sign | Pay Component |
|---:|---|---:|---|
| 1000.02 | FSA HC ISolved | -1 | EE - Healthcare FSA |
| 1000.03 | FSA DC ISolved | -1 | EE - Dependent Care FSA |

Pay Component matching is exact after leading and trailing spaces are removed. The value in the lookup should therefore match the source GL wording.

### Duplicate Mapping Rule

The same combination of **Account** and **Pay Component** may appear only once. Duplicate combinations cause the script to stop with:

```text
Duplicate data in Lookup
```

## Source GL Workbook Requirements

Place only one payroll GL workbook in the script folder when running the tool.

The source workbook must contain payroll data in the first three columns:

| Source Column | Use |
|---|---|
| A | Account number and account description |
| B | Pay component or additional payroll detail |
| C | Amount |

The script first attempts to read row 2 as the header row. If that fails, it attempts to use row 1.

### Account Formatting

The script determines the account using:

- The text before the first colon, when a colon exists; or
- The first eight characters of the value when no colon exists

Example:

```text
1000.01:EE and ER Taxes
```

is mapped to:

```text
1000.01
```

### Required Bank Debit Line

The script searches column A for this exact value:

```text
10012.01:JP Morgan
```

The amount in column C on that line is used as the bank debit for reconciliation. Source rows after the first matching bank line are not included in the journal entry.

### Excluded Payroll Suspense Line

The following exact value is excluded from processing:

```text
20600.07:Payroll Suspense
```

These values can be changed in the constants section of the Python script when a different GL structure is required.

## How to Run

### 1. Download or Clone the Repository

Using Git:

```bash
git clone <repository-url>
cd Payroll-JE-Builder
```

The files may also be downloaded as a ZIP and extracted to a local folder.

### 2. Install Python

Download and install Python 3.10 or newer. During Windows installation, select:

```text
Add Python to PATH
```

Verify the installation:

```bash
python --version
```

On some systems, the command may be:

```bash
python3 --version
```

### 3. Install the Required Packages

```bash
pip install pandas numpy openpyxl
```

If `pip` is associated with another Python installation, use:

```bash
python -m pip install pandas numpy openpyxl
```

### 4. Prepare the Lookup File

1. Make a copy of `Lookup_template.xlsx`.
2. Rename the copy to `Lookup.xlsx`.
3. Update the account mappings, memos, signs, and optional pay components.
4. Save and close the workbook.

### 5. Prepare the Working Folder

Place the following files together in one folder:

```text
build_payroll_journal_final_v29.py
Lookup.xlsx
<one payroll GL workbook>
```

Do not leave unrelated Excel workbooks in the folder. The script requires exactly one Excel workbook other than `Lookup.xlsx`.

A recommended payroll-period folder looks like:

```text
Payroll-JE-Builder/
└── 2026-07-10 Payroll/
    ├── build_payroll_journal_final_v29.py
    ├── Lookup.xlsx
    └── CR_General_Ledger_Outbound_Matrix.xlsx
```

Using a separate working folder for each payroll period helps preserve completed workpapers and prevents prior output files from being mistaken for the next input file.

### 6. Close the Excel Files

Close `Lookup.xlsx` and the payroll GL workbook before running the script. Open workbooks may be locked and prevent the script from saving or renaming them.

### 7. Run the Script

Open Command Prompt, PowerShell, Terminal, or the integrated terminal in an editor. Navigate to the working folder:

```bash
cd "C:\Path\To\Payroll-JE-Builder"
```

Run:

```bash
python build_payroll_journal_final_v29.py
```

### 8. Enter the Dates

The script prompts for:

```text
Bank Date for the JE (MM/DD/YYYY):
Employee Paid Date (MM/DD/YYYY):
```

Example:

```text
Bank Date for the JE (MM/DD/YYYY): 07/09/2026
Employee Paid Date (MM/DD/YYYY): 07/10/2026
```

Dates must be entered in `MM/DD/YYYY` format.

## Generated Outputs

### Completed Payroll Workpaper

The original payroll GL workbook is updated with a new worksheet named:

```text
Payroll JE
```

If that worksheet already exists, the script creates a numbered variation such as:

```text
Payroll JE 2
```

The original workbook is then renamed to:

```text
Payroll WS_<Bank Date>_<Employee Paid Date>.xlsx
```

Example:

```text
Payroll WS_07.09.2026_07.10.2026.xlsx
```

### Journal-Entry Import CSV

The script creates:

```text
Payroll Import_<Bank Date>_<Employee Paid Date>.csv
```

Example:

```text
Payroll Import_07.09.2026_07.10.2026.csv
```

The CSV contains these columns:

```text
EXTERNAL ID
DATE
SUBSIDIARY
ACCOUNT (MAIN)
PAYEE
CURRENCY
MEMO (MAIN)
ACCOUNT
AMOUNT
MEMO (LINE)
DEPARTMENT
PP Project
```

## Static Journal-Entry Values

The current script includes these fixed values:

| Field | Value |
|---|---|
| Subsidiary | `69` |
| Main Account | `10012.01` |
| Payee | `ONESOURCE VIRTUAL, INC.` |
| Currency | `USD` |
| Department | `Operating : Current` |
| External ID | `CHCK` plus employee paid date in `MMDDYY` format |
| Main Memo | Employee paid date plus `Payroll` |

These settings are located near the top of the script and should be reviewed before using the tool for another company, subsidiary, bank account, payroll provider, or import configuration.

## Reconciliation and Warning Checks

At the end of processing, the terminal displays:

- The GL workbook name
- The new worksheet name
- The CSV filename
- The flipped journal-entry total
- The bank debit amount
- Whether the totals match
- Any accounts that generated lookup warnings

Example successful result:

```text
---- RESULT ----
   GL workbook: CR_General_Ledger_Outbound_Matrix.xlsx
   New sheet:   Payroll JE
   CSV written: Payroll Import_07.09.2026_07.10.2026.csv
   JE total (flipped): 125,000.00
   Bank debit:         125,000.00
   ✅ Totals match the bank debit line.
```

### Warning: Account Not in Lookup

When an account does not have a usable lookup memo, the journal-entry line includes:

```text
WARNING: ACCT NOT IN LOOKUP
```

For an account configured with pay-component splits, any source rows whose pay component does not match the lookup are placed into a catch-all warning line.

Always review and resolve warning lines before importing the CSV.

## Important File-Safety Behavior

The script changes the source GL workbook directly and then renames it.

It uses `os.replace`, which means an existing file with the same final output filename may be overwritten. To protect source data:

- Run the script from a payroll-period working folder
- Keep an untouched copy of the original GL export when required by policy
- Confirm that a completed output with the same dates does not already exist
- Review the Excel workpaper and CSV before importing the journal entry

## Common Errors

### `Lookup.xlsx not found`

**Cause:** The lookup file is missing or still named `Lookup_template.xlsx`.

**Resolution:** Rename the working copy to exactly:

```text
Lookup.xlsx
```

### `No GL workbook found`

**Cause:** No supported payroll workbook is in the script folder.

**Resolution:** Add one `.xlsx`, `.xlsm`, or `.xls` payroll GL workbook.

### `Multiple GL workbooks found`

**Cause:** More than one Excel workbook exists in the folder besides `Lookup.xlsx`.

**Resolution:** Move old inputs and completed payroll workpapers to another folder so only the current GL workbook remains.

### `Duplicate data in Lookup`

**Cause:** The same account and pay-component combination appears more than once.

**Resolution:** Remove or consolidate the duplicate lookup row.

### `Totals DO NOT match`

Possible causes include:

- An incorrect sign in the lookup
- A missing or incorrect account mapping
- An unmatched pay component
- A missing or incorrectly labeled bank debit line
- Source data placed after the bank debit line
- A payroll line intentionally excluded by the script
- A GL format that differs from the expected columns A through C

Review warning lines, lookup signs, source ordering, and the `10012.01:JP Morgan` bank debit amount.

### Permission or Rename Error

**Cause:** The source workbook is open in Excel, the destination file is locked, or the user does not have write access to the folder.

**Resolution:** Close the workbooks, confirm folder permissions, and run the script again.

## Recommended Review Before Import

Before uploading the generated CSV:

1. Confirm the terminal states that totals match.
2. Open the completed payroll workpaper.
3. Review every line containing `WARNING: ACCT NOT IN LOOKUP`.
4. Confirm the bank date and employee paid date.
5. Confirm the subsidiary, main account, payee, currency, and department.
6. Confirm debit and credit signs.
7. Confirm the CSV total agrees with the approved payroll bank debit.
8. Follow the organization’s journal-entry approval and import controls.

## Customization

The main configurable constants appear near the top of the script:

```python
ORIG_SHEET_NAME = "CR General Ledger Outbound Matr"
STOP_LABEL = "10012.01:JP Morgan"
EXCLUDE_LABEL = "20600.07:Payroll Suspense"
LOOKUP_BASENAME = "Lookup.xlsx"
OUTPUT_SHEET_BASE = "Payroll JE"

SUBSIDIARY_CONST = 69
MAIN_ACCOUNT = 10012.01
PAYEE = "ONESOURCE VIRTUAL, INC."
CURRENCY = "USD"
DEPARTMENT = "Operating : Current"
```

Update these only after confirming the required payroll export and accounting import structure.

## Efficiency Impact

The largest efficiency gain comes from replacing repetitive manual work with a repeatable mapping process. Instead of rebuilding account groupings, signs, memos, pay-component splits, formulas, and import formatting each payroll period, the user maintains the lookup table and reviews exceptions.

Actual time savings depend on payroll size and the prior process. The tool is most valuable when:

- The payroll GL contains many accounts or detail rows
- The same mapping rules are reused each pay period
- Several accounts require pay-component separation
- A consistent workpaper and import format is required
- Manual reconciliation and exception identification previously took significant time

The script does not eliminate review or approval controls. It shifts the work from repetitive preparation to exception review and validation.

## Disclaimer

This project is an accounting workflow automation tool and should be tested with non-production data before use. Users are responsible for reviewing the generated journal entry, validating account mappings, retaining appropriate source documentation, and following their organization’s approval, access, and financial-control requirements.
