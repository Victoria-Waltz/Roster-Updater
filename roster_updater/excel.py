"""Shared Excel formatting and baseline workbook writing."""
from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

from .constants import BASELINE_HEADERS
from .utils import clean


NAVY = "17365D"


BLUE = "5B9BD5"


LIGHT_BLUE = "D9EAF7"


LIGHT_GREEN = "E2F0D9"


LIGHT_YELLOW = "FFF2CC"


LIGHT_PINK = "FCE4D6"


LIGHT_RED = "F4CCCC"


LIGHT_GRAY = "E7E6E6"


WHITE = "FFFFFF"


DARK_TEXT = "1F1F1F"


THIN_GRAY = Side(style="thin", color="D9E1F2")


def _safe_cell_value(value: Any) -> Any:
    """Prefix formula-like text with an apostrophe before placing it in Excel.

    Strings starting with =, +, -, or @ are treated as literal text by this
    writer. Numeric values are left unchanged.
    """
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _write_rows(worksheet, start_row: int, rows: Iterable[Iterable[Any]]) -> int:
    """Write values starting at a one-based row and return the last row written.

    Each value goes through _safe_cell_value(). For no input rows, return
    start_row - 1 so callers can handle an empty table without inventing data.
    """
    current = start_row
    for values in rows:
        for column, value in enumerate(values, start=1):
            worksheet.cell(current, column, _safe_cell_value(value))
        current += 1
    return current - 1


def _style_table_sheet(
    worksheet, header_row: int, last_row: int, last_column: int,
    add_filter: bool = True,
) -> None:
    """Apply shared heading/body styles, freeze panes, and an optional filter.

    Coordinates are one-based. This formats a worksheet region; callers add
    structured Excel Table objects separately when data rows exist. Pass
    add_filter=False for those sheets: each table owns its own AutoFilter.
    """
    for cell in worksheet[header_row]:
        if cell.column <= last_column:
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.font = Font(color=WHITE, bold=True, name="Arial", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=THIN_GRAY)
    for row in worksheet.iter_rows(min_row=header_row + 1, max_row=max(header_row + 1, last_row), max_col=last_column):
        for cell in row:
            cell.font = Font(color=DARK_TEXT, name="Arial", size=10)
            cell.alignment = Alignment(vertical="center")
    worksheet.freeze_panes = worksheet.cell(header_row + 1, 1)
    worksheet.sheet_view.showGridLines = False
    # A structured table supplies its own filter; do not duplicate it on the sheet.
    worksheet.auto_filter.ref = (
        f"A{header_row}:{_column_letter(last_column)}{max(header_row, last_row)}"
        if add_filter else None
    )


def _column_letter(column: int) -> str:
    """Convert a positive one-based column number to Excel letters (1=A, 27=AA)."""
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _fit_columns(worksheet, minimum: int = 11, maximum: int = 34) -> None:
    """Estimate column widths from up to 400 cells per column.

    Clamp widths between minimum and maximum so long notes do not make the
    workbook difficult to navigate. This is a text-length estimate, not autofit.
    """
    for column_cells in worksheet.columns:
        letter = _column_letter(column_cells[0].column)
        length = 0
        for cell in column_cells[: min(len(column_cells), 400)]:
            length = max(length, len(clean(cell.value)))
        worksheet.column_dimensions[letter].width = max(minimum, min(maximum, length + 2))


def write_baseline(path: str | Path, rows: list[dict[str, Any]]) -> Path:
    """Save canonical employee rows to a Current Baseline worksheet.

    Create parent folders, apply standard formatting, and return the output
    Path. Save a temporary workbook beside the destination, then replace the
    destination so readers do not observe a partially written workbook.
    """
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Current Baseline"
    _write_rows(worksheet, 1, [BASELINE_HEADERS])
    last_row = _write_rows(
        worksheet,
        2,
        ([clean(row.get(header)) for header in BASELINE_HEADERS] for row in rows),
    )
    _style_table_sheet(worksheet, 1, last_row, len(BASELINE_HEADERS), add_filter=False)
    _fit_columns(worksheet)
    # Excel structured tables need data rows; an empty roster keeps headings only.
    if last_row >= 2:
        table = Table(displayName="CurrentBaselineTable", ref=f"A1:O{last_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        worksheet.add_table(table)
    # Save beside the destination, then replace it after workbook creation succeeds.
    temp_path = output.with_suffix(output.suffix + ".tmp")
    workbook.save(temp_path)
    os.replace(temp_path, output)
    return output

