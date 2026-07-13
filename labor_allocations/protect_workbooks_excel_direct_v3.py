from pathlib import Path
import tkinter as tk
from tkinter import filedialog

try:
    import win32com.client as win32
except ImportError:
    raise ImportError(
        "Missing required package: pywin32\n\n"
        "Install it by running this in Command Prompt:\n"
        "pip install pywin32"
    )

PASSWORD = "PPLabor"

# Pick folder
root = tk.Tk()
root.withdraw()

folder = filedialog.askdirectory(
    title="Select folder containing Excel workbooks"
)

if not folder:
    raise Exception("No folder selected.")

folder_path = Path(folder)

excel_files = [
    p for p in folder_path.iterdir()
    if p.suffix.lower() in [".xlsx", ".xlsm"]
    and not p.name.startswith("~$")
]

if not excel_files:
    raise Exception("No .xlsx or .xlsm files found in the selected folder.")

excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    for file_path in excel_files:
        print(f"Processing: {file_path.name}")

        wb = excel.Workbooks.Open(str(file_path))

        try:
            for ws in wb.Worksheets:
                print(f"  Sheet: {ws.Name}")

                # Unprotect first in case the sheet is already protected
                try:
                    ws.Unprotect(Password=PASSWORD)
                except Exception:
                    try:
                        ws.Unprotect()
                    except Exception:
                        pass

                # Direct Excel method:
                # 1. Unlock the entire sheet
                # 2. Lock only columns A:G
                # 3. Lock only rows 1:6
                # 4. Protect sheet
                ws.Cells.Locked = False
                ws.Range("A:G").Locked = True
                ws.Rows("1:6").Locked = True

                ws.Protect(
                    Password=PASSWORD,
                    DrawingObjects=True,
                    Contents=True,
                    Scenarios=True,
                    UserInterfaceOnly=False,
                    AllowFormattingCells=False,
                    AllowFormattingColumns=False,
                    AllowFormattingRows=False,
                    AllowInsertingColumns=False,
                    AllowInsertingRows=False,
                    AllowInsertingHyperlinks=False,
                    AllowDeletingColumns=False,
                    AllowDeletingRows=False,
                    AllowSorting=False,
                    AllowFiltering=False,
                    AllowUsingPivotTables=False
                )

                # Allow users to select editable cells only
                ws.EnableSelection = 1  # xlUnlockedCells

            wb.Save()
            print(f"Saved: {file_path.name}")

        finally:
            wb.Close(SaveChanges=True)

finally:
    excel.DisplayAlerts = True
    excel.Quit()

print("Finished. Only columns A:G and rows 1:6 should be locked.")
