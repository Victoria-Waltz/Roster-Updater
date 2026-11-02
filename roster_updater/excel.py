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
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _write_rows(worksheet, start_row: int, rows: Iterable[Iterable[Any]]) -> int:
    current = start_row
    for values in rows:
        for column, value in enumerate(values, start=1):
            worksheet.cell(current, column, _safe_cell_value(value))
        current += 1
    return current - 1


def _style_table_sheet(worksheet, header_row: int, last_row: int, last_column: int) -> None:
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
    worksheet.auto_filter.ref = f"A{header_row}:{_column_letter(last_column)}{max(header_row, last_row)}"


def _column_letter(column: int) -> str:
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _fit_columns(worksheet, minimum: int = 11, maximum: int = 34) -> None:
    for column_cells in worksheet.columns:
        letter = _column_letter(column_cells[0].column)
        length = 0
        for cell in column_cells[: min(len(column_cells), 400)]:
            length = max(length, len(clean(cell.value)))
        worksheet.column_dimensions[letter].width = max(minimum, min(maximum, length + 2))


def write_baseline(path: str | Path, rows: list[dict[str, Any]]) -> Path:
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
    _style_table_sheet(worksheet, 1, last_row, len(BASELINE_HEADERS))
    _fit_columns(worksheet)
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
    temp_path = output.with_suffix(output.suffix + ".tmp")
    workbook.save(temp_path)
    os.replace(temp_path, output)
    return output
