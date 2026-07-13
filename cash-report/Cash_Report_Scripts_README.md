# Cash Report Automation Scripts

Automates the monthly Cash Report workflow by merging detail report data, distributing rows to reporting tabs based on lookup rules, updating checker formulas, refreshing pivots when Excel automation is available, and creating simplified distribution-ready workbook copies.

## GitHub Repository Description

**Cash Report Automation Toolkit** — Python scripts that streamline recurring accounting cash report preparation by merging detail reports, applying lookup-driven reporting logic, refreshing workbook outputs, and generating cleaner SG&A / Development distribution files.

Alternative shorter description:

> Python automation toolkit for monthly accounting cash reports, including detail-report merging, lookup-based distribution, checker updates, pivot refreshes, and distribution-ready workbook formatting.

## What These Scripts Do

This folder contains two main scripts plus a lookup workbook template/example.

### `merge_pp_esr_main_v23.py`

This is the main cash report processing script.

It:

1. Finds the single cash report `.xlsx` workbook in the same folder as the script.
2. Prompts for the first day of the reporting period in `MM/DD/YYYY` format.
3. Reads `Cash_Report_Lookup.xlsx` to determine which detail report sheets should be merged.
4. Merges data from the detail report sheets into `Merge_PP_ESR`.
5. Copies values only from columns `A:Y`.
6. Overwrites column `V` in the merged output with the original detail report sheet name that each row came from.
7. Adds reporting period date fields into columns `Z:AC`.
8. Distributes merged rows into reporting sheets based on the lookup rules.
9. Handles normal sections and special handling sections separately.
10. Prevents rows added to special handling from also remaining in the normal section of that same report sheet.
11. Updates `Checkers!C129:C131` to sum the special handling sections listed in `Checkers!B129:B131`.
12. Saves a new workbook ending in `_reported.xlsx`.
13. Attempts to refresh pivot tables using Excel COM if `pywin32` and desktop Excel are available.

### `cash_report_distribution_formatter_v3.py`

This script creates cleaner distribution copies from a completed cash report workbook.

It:

1. Prompts for the Excel file to reformat.
2. Creates an `SGNA_` copy of the workbook.
3. Creates a `CIP_Dev_` copy of the workbook.
4. Keeps only the required sheets for each distribution copy.
5. Deletes unnecessary columns.
6. Reorders columns by header name.
7. Formats data rows to match row 2.
8. Formats `Date` columns as short dates.
9. Saves two distribution-ready workbooks in the same folder.

## Why This Exists

The cash report process involves repetitive workbook preparation steps that are easy to get wrong manually, especially when columns are added, lookup rules change, or rows need to be split between normal reporting sections and special handling sections.

These scripts help by:

- Reducing manual copy/paste work.
- Keeping the merge process consistent month to month.
- Preserving the source detail report sheet name in the merged data.
- Applying lookup rules consistently.
- Reducing the risk of duplicate rows between normal and special handling sections.
- Automatically updating reporting period fields.
- Creating cleaner distribution files for SG&A and Development / CIP reporting.
- Saving time during recurring monthly accounting close tasks.

## Files in This Folder

Recommended folder contents:

```text
Cash Report Automation/
├─ merge_pp_esr_main_v23.py
├─ cash_report_distribution_formatter_v3.py
├─ Cash_Report_Lookup.xlsx
├─ Cash_Report_Lookup_example.xlsx        # optional example/template
└─ Your_Cash_Report_Workbook.xlsx         # the workbook to process
```

> Important: `merge_pp_esr_main_v23.py` expects the live lookup workbook to be named exactly `Cash_Report_Lookup.xlsx`.

## Requirements

### Python Packages

Install the required packages with:

```bash
pip install openpyxl
```

Optional, for pivot table refresh through desktop Excel:

```bash
pip install pywin32
```

### System Requirements

- Windows is recommended, especially if using pivot refresh.
- Microsoft Excel desktop app is required for the optional pivot refresh step.
- Python 3.x is required.

## Lookup Workbook Requirements

The main script expects a lookup workbook named:

```text
Cash_Report_Lookup.xlsx
```

It should be saved in the same folder as `merge_pp_esr_main_v23.py`.

### Required Tabs

The lookup workbook must include:

1. `lookup`
2. `Bank_Detail_Reports`

### `Bank_Detail_Reports` Tab

This tab tells the script which detail report sheets to merge.

Required setup:

| Cell / Column | Requirement |
|---|---|
| `A1` | Must equal `Sheet Name` |
| `A2:A` | List the detail report sheet names to merge |

The script reads this list and merges matching sheets from the cash report workbook into `Merge_PP_ESR`.

### `lookup` Tab

This tab controls how rows are distributed from `Merge_PP_ESR` into the reporting sheets.

Expected columns:

| Column | Header | Purpose |
|---|---|---|
| A | `Reporting_Sheets` | Target report sheet name |
| B | `GL Sub-Type` | Include matches based on shifted source column `R` |
| C | `GL Include` | Include exact matches based on source column `G` |
| D | `GL Include 2` | Include based on first 3 characters of source column `G` |
| E | `GL Exclude` | Exclude based on first 3 characters of source column `G` |
| F | `Special Handling` | Sends matching rows to the special handling section |

Multiple lookup values should be separated with `<`.

Example:

```text
DEVELOPMENT<CONSTRUCTION<Pre-NTP
```

## Workbook Requirements

The cash report workbook should contain the expected sheets used by the script.

Common required or expected sheets include:

```text
Merge_PP_ESR
Checkers
PP_Detail_Report
PPM_Detail_Report
PPS_CPB_Detail_Report
PPS_Detail_Report
ESR_Detail_Report
SG&A_Exclusion_Items
PP_Detail_PP_SGA
Dev_Exclusion_Items
PP_Detail_PP_Dev
PP_Detail_PP_Dev_&_CIP
PP_Detail_PP_Dev_CIP_Pivot
PP_SG&A_Pivots
PP_Dev_Pivots
```

The exact detail report sheets merged are controlled by `Cash_Report_Lookup.xlsx` → `Bank_Detail_Reports`.

## How to Run the Main Cash Report Script

1. Put `merge_pp_esr_main_v23.py` in the same folder as the cash report workbook.
2. Put `Cash_Report_Lookup.xlsx` in that same folder.
3. Make sure there is only one cash report `.xlsx` workbook in the folder besides `Cash_Report_Lookup.xlsx`.
4. Open Command Prompt or PowerShell.
5. Navigate to the script folder.

Example:

```bash
cd "C:\Users\YourName\Desktop\Cash Report Automation"
```

6. Run the script:

```bash
python merge_pp_esr_main_v23.py
```

7. Enter the first day of the reporting period when prompted.

Example:

```text
Enter the first day of the reporting period (MM/DD/YYYY): 07/01/2026
```

8. The script will create a new workbook with `_reported` added to the file name.

Example output:

```text
Monthly_Cash_Report_reported.xlsx
```

## How to Run the Distribution Formatter Script

Run this after the main reported workbook has been created and reviewed.

1. Put `cash_report_distribution_formatter_v3.py` in the same folder as the completed workbook.
2. Open Command Prompt or PowerShell.
3. Navigate to the script folder.
4. Run the script:

```bash
python cash_report_distribution_formatter_v3.py
```

5. Enter the workbook file name when prompted.

Example:

```text
What is the excel file you would like to reformat for distribution?
Monthly_Cash_Report_reported.xlsx
```

6. The script will create two new files:

```text
SGNA_Monthly_Cash_Report_reported.xlsx
CIP_Dev_Monthly_Cash_Report_reported.xlsx
```

## Important Notes

- The scripts should be run from the same folder as the workbook files.
- Do not rename `Cash_Report_Lookup.xlsx` unless you also update the script configuration.
- The main script expects only one process workbook in the folder, excluding `Cash_Report_Lookup.xlsx`.
- The formatter script relies on specific worksheet names and required headers.
- If required headers are missing, the formatter will stop and identify the missing column.
- Pivot refresh is optional. If `pywin32` or Excel COM is unavailable, the script will still save the workbook and skip the pivot refresh step.
- The scripts are designed around the current cash report structure where two columns were inserted after column `J`: `Cash Impact` and `Excluded`.

## Troubleshooting

### Error: More than one Excel file found

The main script only allows one cash report workbook in the folder besides `Cash_Report_Lookup.xlsx`.

Move old files, examples, or backups to another folder and run the script again.

### Error: Lookup file not found

Make sure the lookup file is named exactly:

```text
Cash_Report_Lookup.xlsx
```

and is saved in the same folder as `merge_pp_esr_main_v23.py`.

### Error: Required lookup tab not found

Make sure the lookup workbook includes:

```text
lookup
Bank_Detail_Reports
```

### Warning: Sheet listed in lookup but not found in workbook

A sheet name listed in `Bank_Detail_Reports` does not exist in the cash report workbook. The script will skip that sheet and continue.

Check for spelling, spacing, or renamed detail report sheets.

### Pivot refresh skipped

This usually means `pywin32` is not installed or desktop Excel automation is unavailable.

Install it with:

```bash
pip install pywin32
```

Then rerun the script if pivot refresh is needed.

## Suggested Run Order

```text
1. Prepare the cash report workbook.
2. Confirm Cash_Report_Lookup.xlsx is updated.
3. Run merge_pp_esr_main_v23.py.
4. Review the _reported workbook.
5. Run cash_report_distribution_formatter_v3.py.
6. Review the SGNA_ and CIP_Dev_ distribution files.
```

## Suggested GitHub Repository Structure

```text
Cash-Report-Automation-Toolkit/
├─ README.md
├─ scripts/
│  ├─ merge_pp_esr_main_v23.py
│  └─ cash_report_distribution_formatter_v3.py
├─ examples/
│  └─ Cash_Report_Lookup_example.xlsx
└─ docs/
   └─ workflow_notes.md
```

## Security / Data Reminder

Before publishing this repository publicly, remove or anonymize any workbook that contains company data, vendor data, bank detail data, internal IDs, employee information, or confidential accounting information.

The safest public version should include:

- The Python scripts.
- A sanitized lookup workbook example.
- A README file.
- No real cash report workbooks.
