"""Batch-copy Excel cell values from column G to column I, starting at row 10.

Processes .xlsx files in the same folder as this script (or bundled .exe),
writing copies named <original>_1.xlsx. Requires Microsoft Excel and xlwings.
"""

from __future__ import annotations

import sys
from pathlib import Path

import xlwings as xw

START_ROW = 10
SOURCE_COLUMN = "G"
TARGET_COLUMN = "I"


def working_folder() -> Path:
    """Use the executable's folder when packaged, otherwise the script's folder."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def input_workbooks(folder: Path) -> list[Path]:
    """Avoid Excel lock files and previously generated results."""
    return sorted(
        (file for file in folder.iterdir()
         if file.is_file()
         and file.suffix.lower() == ".xlsx"
         and not file.name.startswith("~$")
         and not file.stem.lower().endswith("_1")),
        key=lambda file: file.name.lower(),
    )


def copy_values(app: xw.App, source: Path) -> str:
    """Create a separate output workbook; never overwrite an existing result."""
    output = source.with_name(f"{source.stem}_1.xlsx")
    if output.exists():
        return f"SKIPPED: {output.name} already exists"

    workbook = None
    try:
        workbook = app.books.open(str(source), update_links=False)
        sheet = workbook.sheets[0]

        # Search upwards in column G, rather than downwards from G10.
        # The old approach could jump to Excel's final row if G10 was empty.
        last_excel_row = sheet.cells.last_cell.row
        last_row = sheet.range(f"{SOURCE_COLUMN}{last_excel_row}").end("up").row
        if last_row < START_ROW:
            return f"SKIPPED: {source.name} has no data in {SOURCE_COLUMN}{START_ROW} onwards"

        source_range = sheet.range(f"{SOURCE_COLUMN}{START_ROW}:{SOURCE_COLUMN}{last_row}")
        target_range = sheet.range(f"{TARGET_COLUMN}{START_ROW}:{TARGET_COLUMN}{last_row}")

        # ndim=2 keeps single-cell and multi-row copies consistently shaped.
        # This copies values, not formulas or formatting, like the original tool.
        target_range.value = source_range.options(ndim=2).value
        workbook.save(str(output))
        return f"CREATED: {output.name} (rows {START_ROW}-{last_row})"
    finally:
        if workbook is not None:
            workbook.close()


def main() -> int:
    folder = working_folder()
    workbooks = input_workbooks(folder)
    print(f"Folder: {folder}")
    if not workbooks:
        print("No input .xlsx files found.")
        return 0

    errors = 0
    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        for workbook in workbooks:
            try:
                print(copy_values(app, workbook))
            except Exception as exc:
                errors += 1
                print(f"ERROR: {workbook.name}: {exc}", file=sys.stderr)

    print(f"Finished: {len(workbooks)} source file(s), {errors} error(s).")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
