"""Read and validate CSV and Excel employee rosters."""
from __future__ import annotations

from pathlib import Path
from typing import Any
import csv

from openpyxl import load_workbook

from .constants import BASELINE_HEADERS, CORE_HEADERS, KEY_HEADER
from .models import LoadedRoster, RosterError
from .utils import clean, file_sha256, normalize_header, normalize_key


def _validate_rows(rows: list[dict[str, Any]], source: Path) -> None:
    """Reject empty rosters and duplicate nonblank Logon Name values.

    Matching is case-insensitive. The caller filters out blank-key rows.
    Raise RosterError with the source filename and up to five duplicate keys.
    """
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
    """Convert one input record to the standard baseline column layout.

    Discard unrecognized columns, normalize recognized values to text, and
    fill missing standard fields with empty strings, including ERP Category.
    """
    canonical: dict[str, Any] = {}
    for header, value in raw.items():
        mapped = normalize_header(header)
        if mapped in BASELINE_HEADERS:
            canonical[mapped] = clean(value)
    for header in BASELINE_HEADERS:
        canonical.setdefault(header, "")
    return canonical


def _find_xlsx_source(path: Path, preferred_sheet: str | None) -> tuple[str, int, list[str]]:
    """Find the most likely header row within the first ten rows of a workbook.

    Search only preferred_sheet when it exists; otherwise search every sheet.
    Score recognized headings, with extra weight for Logon Name. Return the
    sheet name, one-based header row, and normalized headings. Full required
    column validation happens later in load_roster().
    """
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
                # Some exports have title rows above their table; score likely
                # headers rather than assuming row 1 always contains headings.
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
    """Read a CSV, XLSX, or XLSM file into a validated LoadedRoster.

    Args:
        path: Input employee file; expanded and resolved to an absolute path.
        preferred_sheet: Excel sheet to prioritize when present; ignored for CSV.

    Returns:
        Canonical employee rows, source metadata, and a SHA-256 fingerprint.

    Raises:
        RosterError: Missing file, unsupported type, missing required headings,
            empty employee data, or duplicate employee keys. Underlying file
            and workbook errors may also propagate to the caller.

    Rows with blank Logon Name values are skipped. Excel formulas are read
    using their cached values; this function does not calculate formulas.
    """
    source = Path(path).expanduser().resolve()
    if not source.exists():
        raise RosterError(f"File not found: {source}")

    suffix = source.suffix.casefold()
    rows: list[dict[str, Any]] = []
    source_headers: list[str] = []
    sheet_name = "CSV"

    if suffix == ".csv":
        # utf-8-sig accepts Excel-exported CSVs with a byte-order mark (BOM).
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
                # Ignore blank/formatting-only rows that cannot identify an employee.
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
                # Ignore blank/formatting-only rows that cannot identify an employee.
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
