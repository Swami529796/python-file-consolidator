"""Excel File Consolidator (Python version).

Combines every Excel / CSV file in a folder into one master workbook.

  - Validates that each file has the same headers as the first file
  - Adds "Source File" and "Source Sheet" columns for traceability
  - Logs every file (Loaded / Skipped / Error) on a "Log" sheet
  - Never modifies the source files

Usage:
    python consolidate.py sample-data
    python consolidate.py "C:/Reports/June" --output June_Master.xlsx --all-sheets

Author : Swaminathan Narayanan
Repo   : https://github.com/Swami529796/python-file-consolidator
"""
from __future__ import annotations

import argparse
import sys
import time
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv"}
HEADER_FILL = PatternFill("solid", start_color="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
STATUS_COLOURS = {"Loaded": "C6EFCE", "Skipped": "FFEB9C", "Error": "FFC7CE"}


@dataclass
class LogEntry:
    """One line of the audit log: what happened to one sheet of one file."""

    file: str
    sheet: str
    rows_added: int
    status: str  # Loaded | Skipped | Error
    message: str = ""


@dataclass
class Result:
    master: pd.DataFrame
    log: list[LogEntry]

    @property
    def loaded(self) -> int:
        return sum(1 for entry in self.log if entry.status == "Loaded")

    @property
    def skipped(self) -> int:
        return sum(1 for entry in self.log if entry.status != "Loaded")


def list_files(folder: Path, exclude: Path | None = None) -> list[Path]:
    """Supported files in the folder, in name order, ignoring Excel lock files."""
    files = [
        p
        for p in sorted(folder.iterdir(), key=lambda p: p.name.lower())
        if p.is_file()
        and p.suffix.lower() in SUPPORTED_EXTENSIONS
        and not p.name.startswith("~$")
    ]
    if exclude is not None:
        files = [p for p in files if p.resolve() != exclude.resolve()]
    return files


def read_sheets(path: Path, all_sheets: bool) -> list[tuple[str, pd.DataFrame]]:
    """Return (sheet name, data) pairs for one file. A CSV counts as one sheet."""
    if path.suffix.lower() == ".csv":
        return [(path.stem, pd.read_csv(path))]
    sheets = pd.read_excel(path, sheet_name=None if all_sheets else 0)
    if isinstance(sheets, pd.DataFrame):  # single sheet requested
        with pd.ExcelFile(path) as book:
            return [(str(book.sheet_names[0]), sheets)]
    return [(str(name), frame) for name, frame in sheets.items()]


def _normalise(header: object) -> str:
    return str(header).strip().lower()


def check_headers(expected: list[str], found: list[str]) -> str:
    """Return an empty string if the headers match, otherwise the reason."""
    if len(found) != len(expected):
        return f"Header mismatch - expected {len(expected)} columns, found {len(found)}"
    for position, (want, got) in enumerate(zip(expected, found), start=1):
        if _normalise(want) != _normalise(got):
            return f"Header mismatch - column {position}: expected '{want}', found '{got}'"
    return ""


def _align_dates(master: pd.DataFrame) -> pd.DataFrame:
    """CSV files give dates as text and Excel files give real dates.

    Where a column already holds real dates from at least one file, convert the
    text values too, so the column has one consistent type.
    """
    for column in master.columns:
        values = master[column]
        if values.dtype != object:
            continue
        has_real_dates = values.map(lambda v: isinstance(v, (datetime, date))).any()
        if not has_real_dates:
            continue
        converted = pd.to_datetime(values, errors="coerce")
        if converted.notna().sum() == values.notna().sum():  # nothing lost
            master[column] = converted
    return master


def consolidate(folder: Path, all_sheets: bool = False, exclude: Path | None = None) -> Result:
    """Merge every supported file in `folder` and return the data plus a log."""
    frames: list[pd.DataFrame] = []
    log: list[LogEntry] = []
    expected: list[str] | None = None

    for path in list_files(folder, exclude):
        try:
            sheets = read_sheets(path, all_sheets)
        except Exception as error:  # a corrupt file must not stop the run
            log.append(LogEntry(path.name, "", 0, "Error", f"Could not open file: {error}"))
            continue

        for sheet_name, frame in sheets:
            frame = frame.dropna(how="all")
            headers = [str(column) for column in frame.columns]
            if frame.empty and not headers:
                log.append(LogEntry(path.name, sheet_name, 0, "Skipped", "Sheet is empty"))
                continue
            if expected is None:
                expected = headers  # the first valid sheet sets the standard
            problem = check_headers(expected, headers)
            if problem:
                log.append(LogEntry(path.name, sheet_name, 0, "Skipped", problem))
                continue

            frame.columns = expected  # same spelling and case everywhere
            frame = frame.assign(**{"Source File": path.name, "Source Sheet": sheet_name})
            frames.append(frame)
            log.append(LogEntry(path.name, sheet_name, len(frame), "Loaded"))

    if frames:
        master = _align_dates(pd.concat(frames, ignore_index=True))
    else:
        master = pd.DataFrame()
    return Result(master, log)


def _style_sheet(sheet) -> None:
    """Bold coloured header, frozen top row, filter and sensible column widths."""
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for index, column in enumerate(sheet.columns, start=1):
        longest = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column)
        sheet.column_dimensions[get_column_letter(index)].width = min(max(longest + 2, 10), 60)


def write_output(result: Result, output: Path) -> None:
    """Write the Master and Log sheets to one formatted workbook."""
    log_frame = pd.DataFrame([asdict(entry) for entry in result.log])
    log_frame.columns = ["File", "Sheet", "Rows Added", "Status", "Message"]

    with pd.ExcelWriter(output, engine="openpyxl", date_format="yyyy-mm-dd",
                        datetime_format="yyyy-mm-dd") as writer:
        result.master.to_excel(writer, sheet_name="Master", index=False)
        log_frame.to_excel(writer, sheet_name="Log", index=False)
        _style_sheet(writer.sheets["Master"])
        log_sheet = writer.sheets["Log"]
        _style_sheet(log_sheet)
        for row in log_sheet.iter_rows(min_row=2, min_col=4, max_col=4):
            colour = STATUS_COLOURS.get(str(row[0].value))
            if colour:
                row[0].fill = PatternFill("solid", start_color=colour)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Merge every Excel/CSV file in a folder.")
    parser.add_argument("folder", type=Path, help="folder containing the files to merge")
    parser.add_argument("-o", "--output", type=Path, default=Path("Master.xlsx"),
                        help="output workbook (default: Master.xlsx)")
    parser.add_argument("--all-sheets", action="store_true",
                        help="read every sheet of each workbook, not just the first")
    args = parser.parse_args(argv)

    if not args.folder.is_dir():
        print(f"Folder not found: {args.folder}", file=sys.stderr)
        return 1

    started = time.perf_counter()
    result = consolidate(args.folder, all_sheets=args.all_sheets, exclude=args.output)
    if not result.log:
        print(f"No Excel or CSV files found in: {args.folder}", file=sys.stderr)
        return 1
    write_output(result, args.output)

    print("Consolidation complete.")
    print(f"  Files found:      {len({entry.file for entry in result.log})}")
    print(f"  Sheets loaded:    {result.loaded}")
    print(f"  Skipped / errors: {result.skipped}")
    print(f"  Rows added:       {len(result.master):,}")
    print(f"  Time taken:       {time.perf_counter() - started:.1f} sec")
    print(f"  Output:           {args.output}")
    for entry in result.log:
        if entry.status != "Loaded":
            print(f"  ! {entry.file}: {entry.message}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
