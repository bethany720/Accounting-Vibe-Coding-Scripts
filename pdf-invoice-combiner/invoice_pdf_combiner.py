"""
Invoice ZIP PDF Combiner

What it does:
1. Accepts a ZIP file path by drag-and-drop, command line, or prompt.
2. Extracts the ZIP to a temporary working folder.
3. Converts Excel files to PDF using Microsoft Excel.
4. Combines PDFs into one final PDF.
5. Prioritizes files with "invoice" or invoice-like "inv" in the filename as the first pages.
6. Saves the combined PDF to an Output folder.

Requirements:
- Windows
- Microsoft Excel installed
- Python packages:
    pip install pypdf pywin32
"""

from __future__ import annotations

import argparse
import datetime as dt
import logging
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable, List, Tuple

from pypdf import PdfWriter


SUPPORTED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
SUPPORTED_PDF_EXTENSIONS = {".pdf"}

# Files/folders inside zips that should generally be ignored
IGNORED_NAME_PARTS = {
    "__macosx",
}

OUTPUT_FOLDER_NAME = "Output"
LOG_FOLDER_NAME = "Logs"


def setup_logging(app_folder: Path) -> Path:
    log_folder = app_folder / LOG_FOLDER_NAME
    log_folder.mkdir(parents=True, exist_ok=True)

    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = log_folder / f"invoice_pdf_combiner_{timestamp}.log"

    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )

    return log_path


def clean_filename_stem(name: str) -> str:
    """
    Make a safe output name based on the ZIP file name.
    """
    cleaned = re.sub(r"[^\w\-. ]+", "_", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or "Combined_File"


def is_ignored_path(path: Path) -> bool:
    parts_lower = [part.lower() for part in path.parts]
    return any(ignored in parts_lower for ignored in IGNORED_NAME_PARTS)


def looks_like_invoice(path: Path) -> bool:
    """
    Decide whether a file should be prioritized as the first page.

    Strong match:
    - invoice anywhere in the file name

    Safer invoice abbreviation matches:
    - inv at the start
    - inv separated by spaces, underscores, dashes, dots, or numbers
    - examples: inv_123.pdf, inv-123.pdf, INV 123.xlsx, 123_inv.pdf
    """
    stem = path.stem.lower()

    if "invoice" in stem:
        return True

    # Match "inv" as a token, not as random letters inside a word.
    # Examples matched: inv_123, inv-123, inv 123, 123_inv, INV.123
    return bool(re.search(r"(^|[\s_\-.#])inv($|[\s_\-.#0-9])", stem))


def sort_key_for_merge(path: Path) -> Tuple[int, str]:
    """
    Invoice-looking files first, then everything else alphabetically.
    """
    priority = 0 if looks_like_invoice(path) else 1
    return priority, path.name.lower()


def extract_zip(zip_path: Path, work_folder: Path) -> Path:
    extract_folder = work_folder / "extracted"
    extract_folder.mkdir(parents=True, exist_ok=True)

    logging.info("Extracting ZIP: %s", zip_path)
    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        zip_ref.extractall(extract_folder)

    return extract_folder


def find_supported_files(folder: Path) -> List[Path]:
    files: List[Path] = []

    for path in folder.rglob("*"):
        if not path.is_file():
            continue

        if is_ignored_path(path):
            continue

        suffix = path.suffix.lower()

        if suffix in SUPPORTED_PDF_EXTENSIONS or suffix in SUPPORTED_EXCEL_EXTENSIONS:
            files.append(path)

    return files


def convert_excel_files_to_pdf(excel_files: Iterable[Path], converted_folder: Path) -> List[Path]:
    """
    Converts Excel files to PDFs using Microsoft Excel via pywin32.

    Notes:
    - This requires Excel to be installed.
    - It exports the active/printable workbook to PDF.
    - If a workbook has bad print areas or huge sheets, the resulting PDF will reflect that.
    """
    excel_files = list(excel_files)
    if not excel_files:
        return []

    converted_folder.mkdir(parents=True, exist_ok=True)

    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is not installed. Run: pip install pywin32"
        ) from exc

    logging.info("Starting Excel for conversion.")
    excel_app = win32com.client.DispatchEx("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    converted_pdfs: List[Path] = []

    try:
        for excel_path in excel_files:
            logging.info("Converting Excel to PDF: %s", excel_path)

            output_pdf = converted_folder / f"{excel_path.stem}.pdf"

            # Avoid overwriting if two files have same stem
            counter = 2
            while output_pdf.exists():
                output_pdf = converted_folder / f"{excel_path.stem}_{counter}.pdf"
                counter += 1

            workbook = None
            try:
                workbook = excel_app.Workbooks.Open(str(excel_path.resolve()))
                workbook.ExportAsFixedFormat(0, str(output_pdf.resolve()))
                converted_pdfs.append(output_pdf)
                logging.info("Converted to PDF: %s", output_pdf)
            except Exception:
                logging.exception("Failed converting Excel file: %s", excel_path)
                raise
            finally:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)

    finally:
        excel_app.Quit()
        logging.info("Excel closed.")

    return converted_pdfs


def merge_pdfs(pdf_files: Iterable[Path], output_pdf: Path) -> None:
    pdf_files = list(pdf_files)

    if not pdf_files:
        raise RuntimeError("No PDF files were found to combine.")

    writer = PdfWriter()

    for pdf_path in pdf_files:
        logging.info("Adding PDF to merge: %s", pdf_path)
        try:
            writer.append(str(pdf_path.resolve()))
        except Exception:
            logging.exception("Failed adding PDF to merge: %s", pdf_path)
            raise

    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    with output_pdf.open("wb") as f:
        writer.write(f)

    writer.close()
    logging.info("Saved combined PDF: %s", output_pdf)


def process_zip(zip_path: Path, app_folder: Path) -> Path:
    if not zip_path.exists():
        raise FileNotFoundError(f"ZIP file does not exist: {zip_path}")

    if zip_path.suffix.lower() != ".zip":
        raise ValueError(f"Input file must be a .zip file: {zip_path}")

    # Save final combined PDFs to the user's Downloads folder.
    # Falls back to a local Output folder if Downloads cannot be found.
    downloads_folder = Path.home() / "Downloads"
    output_folder = downloads_folder if downloads_folder.exists() else app_folder / OUTPUT_FOLDER_NAME
    output_folder.mkdir(parents=True, exist_ok=True)

    safe_stem = clean_filename_stem(zip_path.stem)
    output_pdf = output_folder / f"{safe_stem}_Combined.pdf"

    # Avoid overwriting an existing combined file
    counter = 2
    while output_pdf.exists():
        output_pdf = output_folder / f"{safe_stem}_Combined_{counter}.pdf"
        counter += 1

    with tempfile.TemporaryDirectory(prefix="invoice_pdf_combiner_") as temp_dir:
        work_folder = Path(temp_dir)
        extracted_folder = extract_zip(zip_path, work_folder)

        supported_files = find_supported_files(extracted_folder)
        logging.info("Supported files found: %s", len(supported_files))

        original_pdfs = [p for p in supported_files if p.suffix.lower() in SUPPORTED_PDF_EXTENSIONS]
        excel_files = [p for p in supported_files if p.suffix.lower() in SUPPORTED_EXCEL_EXTENSIONS]

        converted_folder = work_folder / "converted_excel_pdfs"
        converted_pdfs = convert_excel_files_to_pdf(excel_files, converted_folder)

        all_pdfs = original_pdfs + converted_pdfs
        all_pdfs_sorted = sorted(all_pdfs, key=sort_key_for_merge)

        logging.info("Final merge order:")
        for index, pdf in enumerate(all_pdfs_sorted, start=1):
            logging.info("  %s. %s", index, pdf)

        merge_pdfs(all_pdfs_sorted, output_pdf)

    return output_pdf


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Unzip and combine PDF/Excel files into one PDF.")
    parser.add_argument(
        "zip_path",
        nargs="?",
        help="Path to the ZIP file. You can also drag a ZIP file onto this script/exe.",
    )
    return parser.parse_args()


def main() -> int:
    app_folder = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    log_path = setup_logging(app_folder)

    try:
        args = parse_args()

        if args.zip_path:
            zip_path = Path(args.zip_path)
        else:
            entered = input("Drag/paste the ZIP file path here, then press Enter: ").strip().strip('"')
            zip_path = Path(entered)

        output_pdf = process_zip(zip_path, app_folder)

        print()
        print("Success!")
        print(f"Combined PDF saved here:")
        print(output_pdf)
        print()
        print(f"Log saved here:")
        print(log_path)

        # Keep command window open when double-clicked
        input("Press Enter to close...")

        return 0

    except Exception as exc:
        logging.exception("Process failed.")
        print()
        print("Something went wrong.")
        print(str(exc))
        print()
        print(f"Check the log file for details:")
        print(log_path)
        input("Press Enter to close...")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
