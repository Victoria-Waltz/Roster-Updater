"""Build monthly Excel reports and process file-based requests."""
from __future__ import annotations

from collections import Counter
from pathlib import Path
import os

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

from .comparison import compare_rosters
from .constants import BASELINE_HEADERS, CAPABILITY_HEADER, KEY_HEADER
from .corrections import load_overrides
from .excel import (
    LIGHT_GREEN,
    LIGHT_PINK,
    LIGHT_RED,
    LIGHT_YELLOW,
    NAVY,
    WHITE,
    _fit_columns,
    _safe_cell_value,
    _style_table_sheet,
    _write_rows,
)
from .loading import load_roster
from .models import ComparisonResult
from .utils import clean, normalize_key


def write_report(path: str | Path, result: ComparisonResult) -> Path:
    """Export a comparison to Summary, Review, Changes, and Updated Roster sheets.

    Args:
        path: Destination XLSX filename; parent directories are created.
        result: Completed in-memory comparison, including saved corrections
            that were present when the comparison ran.

    Returns:
        The output Path after saving through a temporary file and replacement.

    Colors identify review categories and capability sources. Exporting does
    not change the shared baseline or mark any review item as approved.
    """
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Summary"
    review_sheet = workbook.create_sheet("Review")
    changes_sheet = workbook.create_sheet("Changes")
    roster_sheet = workbook.create_sheet("Updated Roster")

    # Summary: overall metrics plus capability/department counts for quick review.
    summary_sheet.sheet_view.showGridLines = False
    summary_sheet["A2"] = "Monthly employee roster update"
    summary_sheet["A2"].font = Font(name="Arial", size=15, bold=True, color=NAVY)
    summary_sheet["A3"] = f"Created {result.created_at:%B %d, %Y at %I:%M %p}"
    summary_sheet["A3"].font = Font(name="Arial", size=10, italic=True, color="666666")
    summary_sheet["A5"] = "Metric"
    summary_sheet["B5"] = "Count"
    last_summary_row = _write_rows(summary_sheet, 6, result.summary.items())
    _style_table_sheet(summary_sheet, 5, last_summary_row, 2)
    summary_sheet["D5"] = "Source"
    summary_sheet["E5"] = "File"
    summary_sheet["D6"] = "Baseline"
    summary_sheet["E6"] = result.baseline.source_path.name
    summary_sheet["D7"] = "New pull"
    summary_sheet["E7"] = result.monthly_pull.source_path.name
    for cell in summary_sheet[5][3:5]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(color=WHITE, bold=True, name="Arial", size=10)
        cell.alignment = Alignment(horizontal="center")

    capability_counts = Counter(clean(row.get(CAPABILITY_HEADER)) or "Missing" for row in result.updated_roster)
    summary_sheet["D10"] = "Capability"
    summary_sheet["E10"] = "Employees"
    for cell in (summary_sheet["D10"], summary_sheet["E10"]):
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(color=WHITE, bold=True, name="Arial", size=10)
        cell.alignment = Alignment(horizontal="center")
    for row_number, (capability, count) in enumerate(
        sorted(capability_counts.items(), key=lambda item: (-item[1], item[0].casefold())),
        start=11,
    ):
        summary_sheet.cell(row_number, 4, _safe_cell_value(capability))
        summary_sheet.cell(row_number, 5, count)

    department_counts = Counter(clean(row.get("Department")) or "Missing" for row in result.updated_roster)
    summary_sheet["G5"] = "Department"
    summary_sheet["H5"] = "Employees"
    for cell in (summary_sheet["G5"], summary_sheet["H5"]):
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(color=WHITE, bold=True, name="Arial", size=10)
        cell.alignment = Alignment(horizontal="center")
    for row_number, (department, count) in enumerate(
        sorted(department_counts.items(), key=lambda item: (-item[1], item[0].casefold())),
        start=6,
    ):
        summary_sheet.cell(row_number, 7, _safe_cell_value(department))
        summary_sheet.cell(row_number, 8, count)

    summary_sheet.column_dimensions["A"].width = 30
    summary_sheet.column_dimensions["B"].width = 12
    summary_sheet.column_dimensions["D"].width = 18
    summary_sheet.column_dimensions["E"].width = 42
    summary_sheet.column_dimensions["G"].width = 28
    summary_sheet.column_dimensions["H"].width = 12

    # Review: one row per finding, including suggested capabilities and hierarchy.
    # Keep this order aligned with ReviewItem.as_row() in models.py.
    review_headers = [
        "Logon Name",
        "Display Name",
        "Issue",
        "Current Value",
        "Suggested Value",
        "Manager",
        "Manager Email",
        "Manager's Manager",
        "Manager's Manager Email",
        "Changed Fields",
        "Notes",
    ]
    review_sheet["A2"] = "Items to review"
    review_sheet["A2"].font = Font(name="Arial", size=15, bold=True, color=NAVY)
    _write_rows(review_sheet, 4, [review_headers])
    last_review_row = _write_rows(review_sheet, 5, (item.as_row() for item in result.review_items))
    _style_table_sheet(review_sheet, 4, last_review_row, len(review_headers), add_filter=False)
    if last_review_row >= 5:
        table = Table(displayName="ReviewItemsTable", ref=f"A4:K{last_review_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        review_sheet.add_table(table)
        issue_range = f"$C$5:$C${last_review_row}"
        review_sheet.conditional_formatting.add(
            issue_range,
            FormulaRule(formula=['ISNUMBER(SEARCH("Missing from new pull",C5))'], fill=PatternFill("solid", fgColor=LIGHT_RED)),
        )
        review_sheet.conditional_formatting.add(
            issue_range,
            FormulaRule(formula=['ISNUMBER(SEARCH("Capability",C5))'], fill=PatternFill("solid", fgColor=LIGHT_YELLOW)),
        )
        review_sheet.conditional_formatting.add(
            issue_range,
            FormulaRule(formula=['ISNUMBER(SEARCH("Manager",C5))'], fill=PatternFill("solid", fgColor=LIGHT_PINK)),
        )
        review_sheet.conditional_formatting.add(
            issue_range,
            FormulaRule(formula=['ISNUMBER(SEARCH("New employee",C5))'], fill=PatternFill("solid", fgColor=LIGHT_GREEN)),
        )
    _fit_columns(review_sheet, maximum=38)
    review_sheet.column_dimensions["K"].width = 48
    for row_number in range(5, max(5, last_review_row) + 1):
        review_sheet.cell(row_number, 11).alignment = Alignment(vertical="top", wrap_text=True)
        review_sheet.row_dimensions[row_number].height = 32

    # Changes: source field differences, separate from the corrected proposed roster.
    # Keep this order aligned with FieldChange.as_row().
    change_headers = ["Logon Name", "Display Name", "Field", "Old Value", "New Value"]
    changes_sheet["A2"] = "Employee information changes"
    changes_sheet["A2"].font = Font(name="Arial", size=15, bold=True, color=NAVY)
    _write_rows(changes_sheet, 4, [change_headers])
    last_change_row = _write_rows(changes_sheet, 5, (item.as_row() for item in result.changes))
    _style_table_sheet(changes_sheet, 4, last_change_row, len(change_headers), add_filter=False)
    if last_change_row >= 5:
        table = Table(displayName="EmployeeChangesTable", ref=f"A4:E{last_change_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        changes_sheet.add_table(table)
    _fit_columns(changes_sheet, maximum=44)

    # Updated Roster: the proposed active employees plus capability provenance.
    roster_headers = BASELINE_HEADERS + ["Capability Source"]
    _write_rows(roster_sheet, 1, [roster_headers])
    roster_rows = []
    for row in result.updated_roster:
        key = normalize_key(row[KEY_HEADER])
        roster_rows.append([clean(row.get(header)) for header in BASELINE_HEADERS] + [result.capability_sources.get(key, "")])
    last_roster_row = _write_rows(roster_sheet, 2, roster_rows)
    _style_table_sheet(roster_sheet, 1, last_roster_row, len(roster_headers), add_filter=False)
    if last_roster_row >= 2:
        table = Table(displayName="UpdatedRosterTable", ref=f"A1:P{last_roster_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        roster_sheet.add_table(table)
        capability_column = 9
        source_column = 16
        for row_number in range(2, last_roster_row + 1):
            source = clean(roster_sheet.cell(row_number, source_column).value)
            capability_cell = roster_sheet.cell(row_number, capability_column)
            if source in {"Baseline", "Manager suggestion"}:
                capability_cell.fill = PatternFill("solid", fgColor=LIGHT_YELLOW)
            elif source in {"Slack/manual", "Manual", "Other"}:
                capability_cell.fill = PatternFill("solid", fgColor=LIGHT_PINK)
            elif source == "Missing":
                capability_cell.fill = PatternFill("solid", fgColor=LIGHT_RED)
    _fit_columns(roster_sheet, maximum=36)

    # Fit printed sheets to one page wide while allowing multiple pages vertically.
    for worksheet in workbook.worksheets:
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0
        worksheet.sheet_properties.outlinePr.summaryBelow = True

    temp_path = output.with_suffix(output.suffix + ".tmp")
    workbook.save(temp_path)
    os.replace(temp_path, output)
    return output


def create_report_from_files(
    baseline_path: str | Path,
    monthly_path: str | Path,
    report_path: str | Path,
    overrides_path: str | Path | None = None,
) -> ComparisonResult:
    """Load two employee files, compare them, and save an Excel report.

    Use this convenience function from scripts that do not need the desktop
    window or SharedRosterStore. overrides_path optionally supplies saved
    corrections. Return the ComparisonResult for further inspection; this
    function does not approve or replace a shared baseline.
    """
    baseline = load_roster(baseline_path, preferred_sheet="All_Users_Active_Info (1)")
    monthly = load_roster(monthly_path)
    overrides = load_overrides(Path(overrides_path)) if overrides_path else {}
    result = compare_rosters(baseline, monthly, overrides)
    write_report(report_path, result)
    return result

