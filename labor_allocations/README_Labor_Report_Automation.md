# Labor Report Automation Toolkit

A set of Python tools that prepares quarterly labor workpapers, reshapes employee project-allocation data into a NetSuite-ready import format, and applies final worksheet protection across completed Excel workbooks.

## What This Toolkit Does

This repository contains three scripts that support the quarterly labor-report workflow:

1. **Prepare Quarter Workpapers** – renames prior-period workpapers for the new reporting period, copies updated project data and formulas from the `ppprojects` source workbook, carries over Excel formatting, and unprotects worksheets for editing.
2. **Reshape Labor Reports** – converts each department allocation matrix into a long-format `NS Import_Time` worksheet, applies exclusion rules, creates special-project worksheets, and formats the finished output for review or import.
3. **Protect Completed Workbooks** – applies consistent worksheet protection after review is complete, locking columns A:G and rows 1:6 while leaving the remaining allocation area editable.

Together, these scripts reduce repetitive copying, renaming, formatting, exclusion review, and workbook-protection work during each quarterly labor-report cycle.

## Why It Was Built

Quarterly labor reporting can require many nearly identical Excel workpapers to be renamed, updated, reviewed, reshaped, and protected. Completing those steps manually is time-consuming and creates opportunities for inconsistent formatting, missed exclusions, formula errors, and accidental changes to source information.

This toolkit was built to:

- Standardize the quarterly workpaper setup process.
- Reuse one approved project source across every department workbook.
- Preserve formulas, formatting, highlighting, row heights, and column widths.
- Convert wide allocation matrices into a consistent import-ready structure.
- Apply project-level and employee-level exclusion rules automatically.
- Route designated Project IDs to special review worksheets.
- Reduce repetitive manual work and improve consistency between reporting periods.
- Apply the same final protection settings to every completed workbook.

## Included Files

| File | Purpose |
|---|---|
| `Prepare_Quarter_Workpapers_V16.py` | Renames workpapers, copies source values/formulas, pastes Excel formatting, and unprotects worksheets. |
| `reshape_all_in_folder_v4.py` | Creates the `NS Import_Time` output and applies exclusions, sorting, formatting, and special-sheet rules. |
| `protect_workbooks_excel_direct_v3.py` | Protects completed workbooks and limits editing to the intended cells. |
| `Exclusions_template.xlsx` | Template for project exclusions, employee exclusions, special-sheet routing, and workpaper filename lookup values. |

## Requirements

### Operating System

- **Windows**
- Microsoft Excel desktop application installed

The preparation and protection scripts use Excel through Windows COM automation and will not run as designed on macOS or Linux.

### Python

Python 3.10 or later is recommended.

Install the required packages from Command Prompt:

```bash
pip install pandas openpyxl pywin32
```

## Recommended Folder Structure

Keep the scripts and `Exclusions.xlsx` in the same main folder. Store the quarterly Excel workpapers in a subfolder.

```text
Labor-Report-Automation/
│
├── Prepare_Quarter_Workpapers_V16.py
├── reshape_all_in_folder_v4.py
├── protect_workbooks_excel_direct_v3.py
├── Exclusions.xlsx
│
└── workpaper_files/
    ├── ppprojects_Q2_2026.xlsx
    ├── EE Project Labor_Legal_Q1 2026.xlsx
    ├── EE Project Labor_Accounting_Q1 2026.xlsx
    └── other department workpapers.xlsx
```

Before running the scripts, rename `Exclusions_template.xlsx` to:

```text
Exclusions.xlsx
```

The input subfolder may have another name, but you must enter that exact folder name when prompted.

## Exclusions.xlsx Setup

The workbook contains four control sheets.

### `PP Project`

Used to exclude labor allocations based on the project name.

| Column | Required Value |
|---|---|
| A | PP Project name |
| B | Exclusion reason |

Matching is case-insensitive.

### `Employee Name`

Used to exclude allocations based on employee number.

| Column | Required Value |
|---|---|
| A | Employee name or reference label |
| B | Exclusion reason |
| C | Employee Number |

The reshape script uses the employee number as the matching key.

### `Special Sheets`

Used to copy selected allocation rows into additional review worksheets.

| Column | Required Value |
|---|---|
| A | Destination worksheet name |
| B | ProjectID |

When a matching ProjectID is found:

- The row is highlighted yellow on `NS Import_Time`.
- The row is copied to the named special worksheet.
- Yellow rows are moved to the bottom of `NS Import_Time`.

A ProjectID may be assigned to more than one special worksheet.

### `Worksheets`

Used by the preparation script to identify the department or workpaper name within each filename.

| Column | Required Value |
|---|---|
| A | Filename lookup value, such as `Legal`, `Accounting`, or another department name |

The preparation script finds this value within the existing filename and replaces the remaining period text with the newly entered reporting period.

## Expected Workpaper Layout

The scripts expect each labor workbook to contain a worksheet named:

```text
Department
```

The reshape script expects the following source layout:

| Excel Row | Purpose |
|---|---|
| Row 1 | COD Date |
| Row 2 | Project termination date |
| Row 3 | Project Notes |
| Row 4 | Project Stage |
| Row 5 | ProjectID |
| Row 6 | Column headers and project names |
| Row 7 onward | Employee allocation data |

The script treats columns A:F as employee or department metadata, skips columns G:H, and begins reading project allocation columns from column I onward.

Expected metadata headers include fields such as:

- Vendor
- Contractor
- Termination Date
- Employee Name
- Employee number
- Department

Header spelling matters for fields that are renamed or used directly by the script.

## How to Run the Full Workflow

Close all Excel workbooks before running each script. This is especially important for files that will be opened through Excel automation.

### Step 1: Prepare the New Quarter Workpapers

Run:

```bash
python Prepare_Quarter_Workpapers_V16.py
```

The script will prompt for:

1. The workpaper subfolder name.
2. The new reporting period, such as `Q2 2026`.
3. Final confirmation by typing `YES`.

The script then:

- Finds exactly one workbook whose filename starts with `ppprojects`.
- Reads filename lookup values from `Exclusions.xlsx` → `Worksheets`.
- Displays a rename preview.
- Renames each matched workpaper for the new period.
- Reads the source `Workpaper` sheet from the `ppprojects` workbook.
- Updates the `Department` sheet in each target workbook.
- Copies column G and columns J through the calculated ending column.
- Copies values, formulas, comments, hyperlinks, row heights, and column widths.
- Uses Excel Paste Special to transfer formatting and highlighting.
- Unprotects all worksheets using the configured password.

#### Important Source Requirements

The `ppprojects` workbook must contain a sheet named:

```text
Workpaper
```

Each target workbook must contain a sheet named:

```text
Department
```

The script calculates the copy range once using row 6 of the source sheet, adds 100 columns beyond the last used or formatted column, and adds 50 rows beyond the source worksheet's current maximum row.

### Step 2: Reshape the Prepared Workpapers

Run:

```bash
python reshape_all_in_folder_v4.py
```

The script will prompt for:

1. The subfolder containing the prepared workbooks.
2. The reporting start date in `MM/DD/YYYY` format.
3. The reporting end date in `MM/DD/YYYY` format.

Example:

```text
Enter the name of the subfolder containing the Excel files: workpaper_files
Enter Start Date (MM/DD/YYYY): 04/01/2026
Enter End Date (MM/DD/YYYY): 06/30/2026
```

For every `.xlsx` workbook in the selected folder, the script:

- Reads the `Department` worksheet.
- Converts project allocation columns into individual long-format rows.
- Removes blank and zero-percent allocations.
- Converts whole-number percentages when necessary.
- Adds reporting start and end dates.
- Creates or replaces the `NS Import_Time` worksheet.
- Applies exclusion rules in this order:
  1. Project Notes
  2. PP Project lookup
  3. Employee Number lookup
- Sorts included rows before excluded rows.
- Sorts by Employee and then PP Project.
- Formats percentages and dates.
- Auto-sizes columns and adds filters.
- Highlights special ProjectID rows in yellow.
- Copies special rows to their assigned worksheets.
- Moves highlighted rows to the bottom of `NS Import_Time`.

The final `NS Import_Time` columns are:

```text
PP Project
ProjectID
Project Stage
Project Notes
Proj Terminated Date
COD Date
Employee
Percent
Emp Id
External ID
Vendor
Contractor
Emp Termination Date
Department
Start Date
End Date
Exclusions
Reason
```

### Step 3: Review the Results

Before protecting the workbooks, review:

- Allocation percentages.
- Included and excluded rows.
- Exclusion reasons.
- Project IDs and project metadata.
- Yellow-highlighted special-project rows.
- Special worksheets created by the script.
- Reporting start and end dates.
- Formulas and formatting on the `Department` sheet.

Save and close all reviewed workbooks before continuing.

### Step 4: Protect the Completed Workbooks

Run:

```bash
python protect_workbooks_excel_direct_v3.py
```

A folder selection window will appear. Select the folder containing the completed `.xlsx` or `.xlsm` workbooks.

For every worksheet in every selected workbook, the script:

- Attempts to remove existing protection.
- Unlocks the worksheet initially.
- Locks columns A:G.
- Locks rows 1:6.
- Protects the worksheet using the configured password.
- Allows users to select only unlocked cells.
- Saves the workbook in place.

This step should normally be run last, after all data review and corrections are complete.

## Recommended Run Order

```text
1. Prepare_Quarter_Workpapers_V16.py
2. reshape_all_in_folder_v4.py
3. Review completed workbooks
4. protect_workbooks_excel_direct_v3.py
```

## Safety and Backup Recommendations

These scripts modify Excel files directly. Before running them against production workpapers:

- Create a backup copy of the entire input folder.
- Test the workflow on copied workbooks first.
- Close Excel before starting.
- Confirm that only one `ppprojects` source workbook is in the workpaper folder.
- Review the rename preview before typing `YES`.
- Confirm that `Exclusions.xlsx` contains current-quarter rules.
- Do not interrupt Excel while the COM-based scripts are running.

## Common Errors and Troubleshooting

### `Missing required package: pywin32`

Install the package:

```bash
pip install pywin32
```

### `Exclusions.xlsx not found`

Confirm that the workbook is named exactly `Exclusions.xlsx` and is stored in the same folder as the Python scripts.

### `Expected exactly 1 file starting with 'ppprojects'`

The workpaper subfolder must contain exactly one Excel workbook whose filename begins with `ppprojects`.

### `Could not find sheet 'Worksheets'`

Confirm that `Exclusions.xlsx` contains a sheet named `Worksheets` and that lookup values are entered in column A.

### `does not contain a sheet named 'Workpaper'`

The `ppprojects` source workbook must contain the source worksheet named `Workpaper`.

### `does not contain a sheet named 'Department'`

Every target workpaper must contain a worksheet named `Department`.

### Formatting or highlighting did not copy

- Close all source and target workbooks before rerunning the preparation script.
- Confirm that Microsoft Excel desktop is installed.
- Do not click or type in Excel while the script is running.
- Rerun using fresh backup copies if Excel previously repaired workbook styles.

### Invalid date format

Enter dates using exactly:

```text
MM/DD/YYYY
```

### No special worksheets were created

Confirm that:

- `Exclusions.xlsx` contains the `Special Sheets` tab.
- The ProjectID values match the values found in the labor workbooks.
- The destination sheet names are valid Excel worksheet names.

## Configuration Notes

Several settings are defined near the top of the scripts and can be adjusted when needed, including:

- Protection password.
- Source and target worksheet names.
- Number of extra rows and columns copied.
- Excel columns copied from the project source.
- Output worksheet name.
- Required final output column order.

Changes should be tested on backup workbooks before being used in the live reporting process.

## Data Security

The scripts run locally and do not upload workbook data to an external service. However, the repository should contain only code and sanitized templates. Do not commit confidential employee data, payroll information, production workpapers, passwords, or completed labor reports to a public GitHub repository.

## License and Use

This toolkit was created for an internal accounting and labor-reporting workflow. Review company policy and remove confidential information before publishing or adapting it for public use.
