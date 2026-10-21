from __future__ import annotations

import csv
import getpass
import hashlib
import json
import os
import re
import shutil
import time
from collections import Counter
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Iterator

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


CORE_HEADERS = [
    "S.No.",
    "Display Name",
    "First Name",
    "Last Name",
    "Logon Name",
    "Email Address",
    "Title",
    "Department",
    "Capability",
    "Manager",
    "Manager Email Address",
    "Office",
    "City",
    "Employee Type",
]

BASELINE_HEADERS = CORE_HEADERS + ["ERP Category"]
KEY_HEADER = "Logon Name"
CAPABILITY_HEADER = "Capability"
ERP_HEADER = "ERP Category"

COMPARE_FIELDS = [
    "Display Name",
    "First Name",
    "Last Name",
    "Email Address",
    "Title",
    "Department",
    "Capability",
    "Manager",
    "Manager Email Address",
    "Office",
    "City",
    "Employee Type",
]

HEADER_ALIASES = {
    "sno": "S.No.",
    "serialnumber": "S.No.",
    "displayname": "Display Name",
    "firstname": "First Name",
    "lastname": "Last Name",
    "logonname": "Logon Name",
    "loginname": "Logon Name",
    "emailaddress": "Email Address",
    "email": "Email Address",
    "title": "Title",
    "department": "Department",
    "capability": "Capability",
    "competency": "Capability",
    "manager": "Manager",
    "manageremailaddress": "Manager Email Address",
    "manageremail": "Manager Email Address",
    "office": "Office",
    "city": "City",
    "employeetype": "Employee Type",
    "erpcategory": "ERP Category",
}

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


class RosterError(Exception):
    """A user-facing roster processing error."""


@dataclass
class LoadedRoster:
    rows: list[dict[str, Any]]
    source_path: Path
    sheet_name: str
    headers: list[str]
    file_hash: str


@dataclass
class ReviewItem:
    logon_name: str
    display_name: str
    issue: str
    current_value: str = ""
    suggested_value: str = ""
    manager: str = ""
    manager_email: str = ""
    managers_manager: str = ""
    managers_manager_email: str = ""
    changed_fields: str = ""
    notes: str = ""

    def as_row(self) -> list[str]:
        return [
            self.logon_name,
            self.display_name,
            self.issue,
            self.current_value,
            self.suggested_value,
            self.manager,
            self.manager_email,
            self.managers_manager,
            self.managers_manager_email,
            self.changed_fields,
            self.notes,
        ]


@dataclass
class FieldChange:
    logon_name: str
    display_name: str
    field_name: str
    old_value: str
    new_value: str

    def as_row(self) -> list[str]:
        return [
            self.logon_name,
            self.display_name,
            self.field_name,
            self.old_value,
            self.new_value,
        ]


@dataclass
class ComparisonResult:
    baseline: LoadedRoster
    monthly_pull: LoadedRoster
    updated_roster: list[dict[str, Any]]
    review_items: list[ReviewItem]
    changes: list[FieldChange]
    capability_sources: dict[str, str]
    summary: dict[str, int]
    created_at: datetime = field(default_factory=datetime.now)

    @property
    def unresolved_count(self) -> int:
        unresolved = {
            "Missing capability",
            "Capability suggestion",
            "Missing manager information",
            "Manager not in active pull",
        }
        return sum(item.issue in unresolved for item in self.review_items)


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_key(value: Any) -> str:
    return clean(value).casefold()


def normalize_header(value: Any) -> str:
    compact = re.sub(r"[^a-z0-9]", "", clean(value).casefold())
    return HEADER_ALIASES.get(compact, clean(value))


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _validate_rows(rows: list[dict[str, Any]], source: Path) -> None:
    if not rows:
        raise RosterError(f"No employee rows were found in {source.name}.")

    keys: list[str] = []
    for row in rows:
        key = normalize_key(row.get(KEY_HEADER))
        if key:
            keys.append(key)

    if not keys:
        raise RosterError(f"No Logon Name values were found in {source.name}.")

    seen: set[str] = set()
    duplicates: set[str] = set()
    for key in keys:
        if key in seen:
            duplicates.add(key)
        seen.add(key)
    if duplicates:
        example = ", ".join(sorted(duplicates)[:5])
        raise RosterError(
            f"Duplicate Logon Name values were found in {source.name}: {example}. "
            "Fix the duplicates before running the update."
        )


def _canonicalize_row(raw: dict[str, Any]) -> dict[str, Any]:
    canonical: dict[str, Any] = {}
    for header, value in raw.items():
        mapped = normalize_header(header)
        if mapped in BASELINE_HEADERS:
            canonical[mapped] = clean(value)
    for header in BASELINE_HEADERS:
        canonical.setdefault(header, "")
    return canonical


def _find_xlsx_source(path: Path, preferred_sheet: str | None) -> tuple[str, int, list[str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        candidates: list[tuple[int, str, int, list[str]]] = []
        names = [preferred_sheet] if preferred_sheet and preferred_sheet in workbook.sheetnames else workbook.sheetnames
        for sheet_name in names:
            worksheet = workbook[sheet_name]
            for row_number, row in enumerate(
                worksheet.iter_rows(min_row=1, max_row=min(10, worksheet.max_row or 10), values_only=True),
                start=1,
            ):
                headers = [normalize_header(value) for value in row]
                score = sum(header in CORE_HEADERS for header in headers)
                if KEY_HEADER in headers:
                    score += 5
                candidates.append((score, sheet_name, row_number, headers))
        if not candidates:
            raise RosterError(f"No worksheets were found in {path.name}.")
        score, sheet_name, header_row, headers = max(candidates, key=lambda item: item[0])
        if score < 10:
            raise RosterError(
                f"Could not find the employee headers in {path.name}. "
                "The file must include Display Name and Logon Name."
            )
        return sheet_name, header_row, headers
    finally:
        workbook.close()


def load_roster(path: str | Path, preferred_sheet: str | None = None) -> LoadedRoster:
    source = Path(path).expanduser().resolve()
    if not source.exists():
        raise RosterError(f"File not found: {source}")

    suffix = source.suffix.casefold()
    rows: list[dict[str, Any]] = []
    source_headers: list[str] = []
    sheet_name = "CSV"

    if suffix == ".csv":
        with source.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            if not reader.fieldnames:
                raise RosterError(f"No header row was found in {source.name}.")
            source_headers = [normalize_header(header) for header in reader.fieldnames]
            missing = [header for header in CORE_HEADERS if header not in source_headers]
            if missing:
                raise RosterError(
                    f"{source.name} is missing required columns: {', '.join(missing)}."
                )
            for raw in reader:
                row = _canonicalize_row(raw)
                if normalize_key(row[KEY_HEADER]):
                    rows.append(row)
    elif suffix in {".xlsx", ".xlsm"}:
        sheet_name, header_row, detected_headers = _find_xlsx_source(source, preferred_sheet)
        workbook = load_workbook(source, read_only=True, data_only=True)
        try:
            worksheet = workbook[sheet_name]
            raw_headers = [cell.value for cell in next(worksheet.iter_rows(min_row=header_row, max_row=header_row))]
            source_headers = [normalize_header(header) for header in raw_headers]
            missing = [header for header in CORE_HEADERS if header not in source_headers]
            if missing:
                raise RosterError(
                    f"{source.name} is missing required columns: {', '.join(missing)}."
                )
            for values in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
                raw = {
                    raw_headers[index]: values[index] if index < len(values) else ""
                    for index in range(len(raw_headers))
                    if raw_headers[index] is not None
                }
                row = _canonicalize_row(raw)
                if normalize_key(row[KEY_HEADER]):
                    rows.append(row)
        finally:
            workbook.close()
    else:
        raise RosterError("Choose an .xlsx, .xlsm, or .csv file.")

    _validate_rows(rows, source)
    return LoadedRoster(
        rows=rows,
        source_path=source,
        sheet_name=sheet_name,
        headers=source_headers,
        file_hash=file_sha256(source),
    )


def load_overrides(path: Path) -> dict[str, dict[str, str]]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RosterError(f"Could not read saved capability corrections: {exc}") from exc
    if not isinstance(data, dict):
        raise RosterError("The saved capability corrections file is invalid.")
    return {
        normalize_key(key): {
            "capability": clean(value.get("capability")),
            "source": clean(value.get("source")) or "Manual",
            "notes": clean(value.get("notes")),
            "updated_by": clean(value.get("updated_by")),
            "updated_at": clean(value.get("updated_at")),
        }
        for key, value in data.items()
        if isinstance(value, dict) and normalize_key(key)
    }


def _row_map(rows: Iterable[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    return {normalize_key(row[KEY_HEADER]): row for row in rows if normalize_key(row[KEY_HEADER])}


def _build_people_indexes(
    rows: Iterable[dict[str, Any]],
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]], set[str]]:
    by_email: dict[str, dict[str, Any]] = {}
    by_name: dict[str, dict[str, Any]] = {}
    duplicate_names: set[str] = set()
    for row in rows:
        email = normalize_key(row.get("Email Address"))
        name = normalize_key(row.get("Display Name"))
        if email:
            by_email[email] = row
        if name:
            if name in by_name:
                duplicate_names.add(name)
            else:
                by_name[name] = row
    for name in duplicate_names:
        by_name.pop(name, None)
    return by_email, by_name, duplicate_names


def _resolve_manager(
    employee: dict[str, Any],
    by_email: dict[str, dict[str, Any]],
    by_name: dict[str, dict[str, Any]],
) -> dict[str, Any] | None:
    manager_email = normalize_key(employee.get("Manager Email Address"))
    if manager_email and manager_email in by_email:
        return by_email[manager_email]
    manager_name = normalize_key(employee.get("Manager"))
    if manager_name and manager_name in by_name:
        return by_name[manager_name]
    return None


def _manager_hierarchy(
    employee: dict[str, Any],
    by_email: dict[str, dict[str, Any]],
    by_name: dict[str, dict[str, Any]],
) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    manager = _resolve_manager(employee, by_email, by_name)
    managers_manager = _resolve_manager(manager, by_email, by_name) if manager else None
    return manager, managers_manager


def compare_rosters(
    baseline: LoadedRoster,
    monthly_pull: LoadedRoster,
    overrides: dict[str, dict[str, str]] | None = None,
) -> ComparisonResult:
    overrides = overrides or {}
    baseline_map = _row_map(baseline.rows)
    monthly_map = _row_map(monthly_pull.rows)

    updated: list[dict[str, Any]] = []
    capability_sources: dict[str, str] = {}
    changes: list[FieldChange] = []
    review: list[ReviewItem] = []
    changed_people: set[str] = set()

    for source_row in monthly_pull.rows:
        key = normalize_key(source_row[KEY_HEADER])
        prior = baseline_map.get(key)
        final_row = {header: clean(source_row.get(header)) for header in BASELINE_HEADERS}
        override = overrides.get(key, {})
        override_capability = clean(override.get("capability"))
        pull_capability = clean(source_row.get(CAPABILITY_HEADER))
        baseline_capability = clean(prior.get(CAPABILITY_HEADER)) if prior else ""

        if override_capability:
            final_row[CAPABILITY_HEADER] = override_capability
            capability_sources[key] = clean(override.get("source")) or "Manual"
        elif pull_capability:
            final_row[CAPABILITY_HEADER] = pull_capability
            capability_sources[key] = "New pull"
        elif baseline_capability:
            final_row[CAPABILITY_HEADER] = baseline_capability
            capability_sources[key] = "Baseline"
        else:
            final_row[CAPABILITY_HEADER] = ""
            capability_sources[key] = "Missing"

        final_row[ERP_HEADER] = clean(prior.get(ERP_HEADER)) if prior else ""
        updated.append(final_row)

        if prior:
            person_changes: list[str] = []
            for field_name in COMPARE_FIELDS:
                old_value = clean(prior.get(field_name))
                raw_new_value = clean(source_row.get(field_name))
                if field_name == CAPABILITY_HEADER and not raw_new_value:
                    continue
                if normalize_key(old_value) != normalize_key(raw_new_value):
                    changes.append(
                        FieldChange(
                            logon_name=clean(source_row[KEY_HEADER]),
                            display_name=clean(source_row["Display Name"]),
                            field_name=field_name,
                            old_value=old_value,
                            new_value=raw_new_value,
                        )
                    )
                    person_changes.append(field_name)
            if person_changes:
                changed_people.add(key)
                review.append(
                    ReviewItem(
                        logon_name=clean(source_row[KEY_HEADER]),
                        display_name=clean(source_row["Display Name"]),
                        issue="Employee information changed",
                        manager=clean(source_row["Manager"]),
                        manager_email=clean(source_row["Manager Email Address"]),
                        changed_fields=", ".join(person_changes),
                    )
                )
        else:
            review.append(
                ReviewItem(
                    logon_name=clean(source_row[KEY_HEADER]),
                    display_name=clean(source_row["Display Name"]),
                    issue="New employee",
                    current_value=clean(final_row[CAPABILITY_HEADER]),
                    manager=clean(source_row["Manager"]),
                    manager_email=clean(source_row["Manager Email Address"]),
                )
            )

    updated_map = _row_map(updated)
    by_email, by_name, _ = _build_people_indexes(updated)

    capability_suggestions = 0
    missing_capabilities = 0
    manager_issue_people: set[str] = set()

    for employee in updated:
        key = normalize_key(employee[KEY_HEADER])
        manager, managers_manager = _manager_hierarchy(employee, by_email, by_name)
        manager_name = clean(employee.get("Manager"))
        manager_email = clean(employee.get("Manager Email Address"))
        managers_manager_name = clean(managers_manager.get("Display Name")) if managers_manager else ""
        managers_manager_email = clean(managers_manager.get("Email Address")) if managers_manager else ""

        if not manager_name or not manager_email:
            manager_issue_people.add(key)
            missing_parts = []
            if not manager_name:
                missing_parts.append("manager name")
            if not manager_email:
                missing_parts.append("manager email")
            review.append(
                ReviewItem(
                    logon_name=clean(employee[KEY_HEADER]),
                    display_name=clean(employee["Display Name"]),
                    issue="Missing manager information",
                    current_value=", ".join(missing_parts),
                    manager=manager_name,
                    manager_email=manager_email,
                )
            )
        elif manager is None:
            manager_issue_people.add(key)
            review.append(
                ReviewItem(
                    logon_name=clean(employee[KEY_HEADER]),
                    display_name=clean(employee["Display Name"]),
                    issue="Manager not in active pull",
                    current_value=manager_name,
                    manager=manager_name,
                    manager_email=manager_email,
                    notes="The manager may have an inactive or deactivated account.",
                )
            )

        capability = clean(employee.get(CAPABILITY_HEADER))
        if not capability:
            missing_capabilities += 1
            suggested = clean(manager.get(CAPABILITY_HEADER)) if manager else ""
            if suggested:
                capability_suggestions += 1
                review.append(
                    ReviewItem(
                        logon_name=clean(employee[KEY_HEADER]),
                        display_name=clean(employee["Display Name"]),
                        issue="Capability suggestion",
                        current_value="",
                        suggested_value=suggested,
                        manager=manager_name,
                        manager_email=manager_email,
                        managers_manager=managers_manager_name,
                        managers_manager_email=managers_manager_email,
                        notes="Suggested from the active manager's capability. Review before applying.",
                    )
                )
            else:
                review.append(
                    ReviewItem(
                        logon_name=clean(employee[KEY_HEADER]),
                        display_name=clean(employee["Display Name"]),
                        issue="Missing capability",
                        manager=manager_name,
                        manager_email=manager_email,
                        managers_manager=managers_manager_name,
                        managers_manager_email=managers_manager_email,
                    )
                )

    departed_keys = [key for key in baseline_map if key not in monthly_map]
    for key in departed_keys:
        employee = baseline_map[key]
        review.append(
            ReviewItem(
                logon_name=clean(employee[KEY_HEADER]),
                display_name=clean(employee["Display Name"]),
                issue="Missing from new pull",
                current_value=clean(employee[CAPABILITY_HEADER]),
                manager=clean(employee["Manager"]),
                manager_email=clean(employee["Manager Email Address"]),
                notes="Review as a possible departure or inactive account.",
            )
        )

    for item in review:
        employee = updated_map.get(normalize_key(item.logon_name))
        if not employee:
            continue
        manager, managers_manager = _manager_hierarchy(employee, by_email, by_name)
        if not item.manager:
            item.manager = clean(employee.get("Manager"))
        if not item.manager_email:
            item.manager_email = clean(employee.get("Manager Email Address"))
        if managers_manager and not item.managers_manager:
            item.managers_manager = clean(managers_manager.get("Display Name"))
            item.managers_manager_email = clean(managers_manager.get("Email Address"))

    new_keys = set(monthly_map) - set(baseline_map)
    summary = {
        "Baseline employees": len(baseline_map),
        "New pull employees": len(monthly_map),
        "New employees": len(new_keys),
        "Missing from new pull": len(departed_keys),
        "Employees with changes": len(changed_people),
        "Changed fields": len(changes),
        "Missing capabilities": missing_capabilities,
        "Capability suggestions": capability_suggestions,
        "Manager issues": len(manager_issue_people),
        "Review items": len(review),
    }

    review.sort(key=lambda item: (item.issue, item.display_name.casefold(), item.logon_name.casefold()))
    changes.sort(key=lambda item: (item.display_name.casefold(), item.field_name.casefold()))

    return ComparisonResult(
        baseline=baseline,
        monthly_pull=monthly_pull,
        updated_roster=updated,
        review_items=review,
        changes=changes,
        capability_sources=capability_sources,
        summary=summary,
    )


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


def write_report(path: str | Path, result: ComparisonResult) -> Path:
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Summary"
    review_sheet = workbook.create_sheet("Review")
    changes_sheet = workbook.create_sheet("Changes")
    roster_sheet = workbook.create_sheet("Updated Roster")

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
    _style_table_sheet(review_sheet, 4, last_review_row, len(review_headers))
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

    change_headers = ["Logon Name", "Display Name", "Field", "Old Value", "New Value"]
    changes_sheet["A2"] = "Employee information changes"
    changes_sheet["A2"].font = Font(name="Arial", size=15, bold=True, color=NAVY)
    _write_rows(changes_sheet, 4, [change_headers])
    last_change_row = _write_rows(changes_sheet, 5, (item.as_row() for item in result.changes))
    _style_table_sheet(changes_sheet, 4, last_change_row, len(change_headers))
    if last_change_row >= 5:
        table = Table(displayName="EmployeeChangesTable", ref=f"A4:E{last_change_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        changes_sheet.add_table(table)
    _fit_columns(changes_sheet, maximum=44)

    roster_headers = BASELINE_HEADERS + ["Capability Source"]
    _write_rows(roster_sheet, 1, [roster_headers])
    roster_rows = []
    for row in result.updated_roster:
        key = normalize_key(row[KEY_HEADER])
        roster_rows.append([clean(row.get(header)) for header in BASELINE_HEADERS] + [result.capability_sources.get(key, "")])
    last_roster_row = _write_rows(roster_sheet, 2, roster_rows)
    _style_table_sheet(roster_sheet, 1, last_roster_row, len(roster_headers))
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

    for worksheet in workbook.worksheets:
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0
        worksheet.sheet_properties.outlinePr.summaryBelow = True

    temp_path = output.with_suffix(output.suffix + ".tmp")
    workbook.save(temp_path)
    os.replace(temp_path, output)
    return output


@contextmanager
def shared_lock(path: Path, purpose: str, timeout_seconds: int = 0) -> Iterator[None]:
    path.parent.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    while True:
        try:
            descriptor = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
                json.dump(
                    {
                        "purpose": purpose,
                        "user": getpass.getuser(),
                        "created_at": datetime.now().isoformat(timespec="seconds"),
                    },
                    handle,
                )
            break
        except FileExistsError as exc:
            if timeout_seconds <= 0 or time.monotonic() - started >= timeout_seconds:
                raise RosterError(
                    "Another user is updating the shared roster. Wait for them to finish and try again."
                ) from exc
            time.sleep(0.25)
    try:
        yield
    finally:
        try:
            path.unlink()
        except FileNotFoundError:
            pass


class SharedRosterStore:
    def __init__(self, root: str | Path):
        self.root = Path(root).expanduser().resolve()
        self.baseline_path = self.root / "current_baseline.xlsx"
        self.overrides_path = self.root / "capability_corrections.json"
        self.history_dir = self.root / "history"
        self.reports_dir = self.root / "reports"
        self.lock_path = self.root / ".roster_update.lock"
        self.audit_path = self.root / "approval_history.csv"

    def prepare(self) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        self.history_dir.mkdir(parents=True, exist_ok=True)
        self.reports_dir.mkdir(parents=True, exist_ok=True)

    def has_baseline(self) -> bool:
        return self.baseline_path.exists()

    def load_baseline(self) -> LoadedRoster:
        if not self.has_baseline():
            raise RosterError("The shared folder has not been initialized with a baseline yet.")
        return load_roster(self.baseline_path, preferred_sheet="Current Baseline")

    def initialize(self, baseline_source: str | Path) -> LoadedRoster:
        self.prepare()
        source = load_roster(baseline_source, preferred_sheet="All_Users_Active_Info (1)")
        with shared_lock(self.lock_path, "Initialize baseline"):
            if self.baseline_path.exists():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                shutil.copy2(self.baseline_path, self.history_dir / f"baseline_before_reinitialize_{timestamp}.xlsx")
            write_baseline(self.baseline_path, source.rows)
        return self.load_baseline()

    def load_overrides(self) -> dict[str, dict[str, str]]:
        return load_overrides(self.overrides_path)

    def save_override(
        self,
        logon_name: str,
        capability: str,
        source: str,
        notes: str = "",
    ) -> None:
        key = normalize_key(logon_name)
        if not key:
            raise RosterError("A Logon Name is required to save a capability correction.")
        if not clean(capability):
            raise RosterError("Enter a capability before saving the correction.")
        self.prepare()
        with shared_lock(self.lock_path, "Save capability correction", timeout_seconds=3):
            overrides = self.load_overrides()
            overrides[key] = {
                "capability": clean(capability),
                "source": clean(source) or "Manual",
                "notes": clean(notes),
                "updated_by": getpass.getuser(),
                "updated_at": datetime.now().isoformat(timespec="seconds"),
            }
            temp_path = self.overrides_path.with_suffix(".json.tmp")
            temp_path.write_text(json.dumps(overrides, indent=2, ensure_ascii=False), encoding="utf-8")
            os.replace(temp_path, self.overrides_path)

    def remove_override(self, logon_name: str) -> None:
        key = normalize_key(logon_name)
        self.prepare()
        with shared_lock(self.lock_path, "Remove capability correction", timeout_seconds=3):
            overrides = self.load_overrides()
            if key in overrides:
                overrides.pop(key)
                temp_path = self.overrides_path.with_suffix(".json.tmp")
                temp_path.write_text(json.dumps(overrides, indent=2, ensure_ascii=False), encoding="utf-8")
                os.replace(temp_path, self.overrides_path)

    def compare(self, monthly_path: str | Path) -> ComparisonResult:
        baseline = self.load_baseline()
        monthly = load_roster(monthly_path)
        return compare_rosters(baseline, monthly, self.load_overrides())

    def approve(self, result: ComparisonResult) -> tuple[Path, Path]:
        self.prepare()
        with shared_lock(self.lock_path, "Approve monthly baseline"):
            if not self.baseline_path.exists():
                raise RosterError("The shared baseline no longer exists. Run the comparison again.")
            current_hash = file_sha256(self.baseline_path)
            if current_hash != result.baseline.file_hash:
                raise RosterError(
                    "The shared baseline changed after this comparison was run. "
                    "Another user may have approved an update. Run the comparison again before approving."
                )

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = self.history_dir / f"baseline_{timestamp}.xlsx"
            source_copy = self.history_dir / f"source_pull_{timestamp}{result.monthly_pull.source_path.suffix.lower()}"
            report_path = self.reports_dir / f"roster_update_{timestamp}.xlsx"
            shutil.copy2(self.baseline_path, backup_path)
            shutil.copy2(result.monthly_pull.source_path, source_copy)
            write_report(report_path, result)
            write_baseline(self.baseline_path, result.updated_roster)

            new_hash = file_sha256(self.baseline_path)
            is_new = not self.audit_path.exists()
            with self.audit_path.open("a", encoding="utf-8", newline="") as handle:
                writer = csv.writer(handle)
                if is_new:
                    writer.writerow(
                        [
                            "Approved At",
                            "Approved By",
                            "Source File",
                            "Previous Baseline Hash",
                            "New Baseline Hash",
                            "Employee Count",
                            "New Employees",
                            "Missing Employees",
                            "Review Items",
                        ]
                    )
                writer.writerow(
                    [
                        datetime.now().isoformat(timespec="seconds"),
                        getpass.getuser(),
                        result.monthly_pull.source_path.name,
                        result.baseline.file_hash,
                        new_hash,
                        result.summary["New pull employees"],
                        result.summary["New employees"],
                        result.summary["Missing from new pull"],
                        result.summary["Review items"],
                    ]
                )
        return report_path, backup_path


def create_report_from_files(
    baseline_path: str | Path,
    monthly_path: str | Path,
    report_path: str | Path,
    overrides_path: str | Path | None = None,
) -> ComparisonResult:
    baseline = load_roster(baseline_path, preferred_sheet="All_Users_Active_Info (1)")
    monthly = load_roster(monthly_path)
    overrides = load_overrides(Path(overrides_path)) if overrides_path else {}
    result = compare_rosters(baseline, monthly, overrides)
    write_report(report_path, result)
    return result

