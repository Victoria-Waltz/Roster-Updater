"""Carry export-time context through Excel review without adding a visible tab."""
import hashlib
import json

from .constants import BASELINE_HEADERS
from .models import RosterError
from .utils import clean

CONTEXT_SHEET = "_RosterContext"
FORMAT = "TAIT reviewed roster v1"
CHUNK_SIZE = 16000  # Stay below Excel's 32,767-character cell limit.


def add_review_context(workbook, result) -> None:
    """Save a snapshot for import validation and correction reconciliation.

    This very-hidden worksheet survives normal Excel Save As operations. It is
    context, not an editable approval form or cryptographic proof of identity.
    Users still verify and approve the reviewed employee rows in the app.
    """
    data = {
        "baseline_hash": result.baseline.file_hash,
        "corrections_hash": result.corrections_hash,
        "monthly_rows": result.monthly_pull.rows,
        "proposed_rows": result.updated_roster,
        "confirmed_links": result.confirmed_logon_changes,
        "source_name": result.monthly_pull.source_path.name,
    }
    payload = json.dumps(data, ensure_ascii=True)
    sheet = workbook.create_sheet(CONTEXT_SHEET)
    sheet.sheet_state = "veryHidden"
    sheet["A1"] = FORMAT
    sheet["B1"] = hashlib.sha256(payload.encode()).hexdigest()
    for row, start in enumerate(range(0, len(payload), CHUNK_SIZE), start=2):
        sheet.cell(row, 1, "json:" + payload[start:start + CHUNK_SIZE])


def read_review_context(workbook) -> dict:
    """Read and validate the saved snapshot; fail clearly for legacy/raw workbooks."""
    message = "This file is not a supported reviewed-roster export. Export a fresh report from the app, edit Roster Draft, and save it as .xlsx."
    if CONTEXT_SHEET not in workbook.sheetnames or "Roster Draft" not in workbook.sheetnames:
        raise RosterError(message)
    sheet = workbook[CONTEXT_SHEET]
    if sheet["A1"].value != FORMAT:
        raise RosterError(message)
    chunks = [row[0] for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True)]
    if not chunks or any(not isinstance(chunk, str) or not chunk.startswith("json:") for chunk in chunks):
        raise RosterError("The workbook's review context is incomplete. Export a fresh report.")
    payload = "".join(chunk[5:] for chunk in chunks)
    if hashlib.sha256(payload.encode()).hexdigest() != sheet["B1"].value:
        raise RosterError("The workbook's review context changed. Export a fresh report.")
    try:
        data = json.loads(payload)
        if not isinstance(data, dict):
            raise ValueError()
        for name in ("baseline_hash", "source_name"):
            if not isinstance(data.get(name), str) or not data[name]:
                raise ValueError()
        # App exports have a corrections snapshot, including "" when none exist.
        if not isinstance(data.get("corrections_hash"), str):
            raise ValueError()
        links = data.get("confirmed_links")
        if not isinstance(links, dict) or any(not isinstance(k, str) or not isinstance(v, str) for k, v in links.items()):
            raise ValueError()
        for name in ("monthly_rows", "proposed_rows"):
            rows = data.get(name)
            if not isinstance(rows, list) or not rows:
                raise ValueError()
            if any(not isinstance(row, dict) or any(not isinstance(v, str) for v in row.values()) for row in rows):
                raise ValueError()
            # Keep only known roster columns when processing workbook context.
            data[name] = [{header: clean(row.get(header)) for header in BASELINE_HEADERS} for row in rows]
    except (ValueError, TypeError):
        raise RosterError(message) from None
    return data
