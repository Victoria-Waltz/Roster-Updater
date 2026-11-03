"""Preview and commit an explicitly reviewed Excel roster as the next baseline."""
from dataclasses import dataclass
from datetime import datetime
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
import csv
import getpass
import json
import os
import shutil

from openpyxl import load_workbook

from .comparison import compare_rosters
from .constants import BASELINE_HEADERS, CAPABILITY_HEADER, KEY_HEADER
from .excel import write_baseline
from .identity import remap_corrections
from .loading import _validate_rows, load_roster
from .models import FieldChange, LoadedRoster, RosterError
from .review_context import read_review_context
from .utils import clean, file_sha256, normalize_key


@dataclass
class ReviewedImport:
    """The concrete file/state shown to a user before the final import confirmation."""
    reviewed: LoadedRoster
    baseline: LoadedRoster
    corrections_hash: str
    overrides: dict
    confirmed_links: dict[str, str]
    edits: list[FieldChange]
    summary: dict[str, int]
    unresolved_count: int


def prepare_reviewed_import(store, path: str | Path) -> ReviewedImport:
    """Validate a saved app export and calculate exactly what its edits will change.

    Does not write files. The reviewed rows must retain the exported active
    employee keys: correct additions/departures/usernames through a new source
    comparison rather than accidentally importing a filtered/partial worksheet.
    """
    source = Path(path).expanduser().resolve()
    if source.suffix.casefold() not in {".xlsx", ".xlsm"}:
        raise RosterError("Choose the saved Excel report containing Roster Draft (.xlsx or .xlsm).")
    source_hash = file_sha256(source)
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        context = read_review_context(workbook)
    finally:
        workbook.close()
    baseline = store.load_baseline()
    if baseline.file_hash != context["baseline_hash"]:
        raise RosterError("This workbook was exported from a different or older baseline. Run a new comparison and export again.")
    corrections_hash = store._corrections_hash()
    if corrections_hash != context["corrections_hash"]:
        raise RosterError("Saved capability corrections changed since export. Export a fresh comparison before importing.")
    overrides = store.load_overrides()
    reviewed = load_roster(source, preferred_sheet="Roster Draft")
    if reviewed.sheet_name != "Roster Draft":
        raise RosterError("The Roster Draft tab was renamed or removed. Export a fresh report.")
    if any(reviewed.headers.count(header) != 1 for header in BASELINE_HEADERS):
        raise RosterError("Keep each exported roster column exactly once, including ERP Category. Restore the Roster Draft headings before importing.")
    if file_sha256(source) != source_hash or reviewed.file_hash != source_hash:
        raise RosterError("The reviewed workbook changed while loading. Save and select it again.")
    for name in ("monthly_rows", "proposed_rows"):
        _validate_rows(context[name], source)
    # Validate confirmed links against the original source snapshot, even when
    # the reviewer legitimately edited supporting fields (title/manager/etc.).
    monthly = LoadedRoster(context["monthly_rows"], source, "Export snapshot", BASELINE_HEADERS, source_hash)
    proposal = compare_rosters(baseline, monthly, overrides, context["confirmed_links"])
    original = {normalize_key(row[KEY_HEADER]): row for row in context["proposed_rows"]}
    expected = {normalize_key(row[KEY_HEADER]): row for row in proposal.updated_roster}
    if original != expected:
        raise RosterError("The export context does not match this baseline/corrections snapshot. Export again.")
    actual = {normalize_key(row[KEY_HEADER]): row for row in reviewed.rows}
    if actual.keys() != original.keys():
        raise RosterError("Roster Draft must retain every exported employee and Logon Name. Restore removed rows/keys, or correct the source file and run a new comparison.")
    edited_fields = []
    reconciled = remap_corrections(overrides, proposal.confirmed_logon_changes)
    for key, row in actual.items():
        before = original[key]
        for field in BASELINE_HEADERS:
            if clean(row.get(field)) != clean(before.get(field)):
                edited_fields.append(FieldChange(row[KEY_HEADER], row["Display Name"], field,
                                                  clean(before.get(field)), clean(row.get(field))))
        # Excel capability edits are authoritative. Persist a nonblank edit as
        # an approved correction so an older correction cannot silently undo it.
        if clean(row[CAPABILITY_HEADER]) != clean(before[CAPABILITY_HEADER]):
            if clean(row[CAPABILITY_HEADER]):
                reconciled[key] = {
                    "capability": clean(row[CAPABILITY_HEADER]), "source": "Reviewed Excel",
                    "notes": f"Approved from reviewed workbook {source.name}",
                    "updated_by": getpass.getuser(),
                    "updated_at": datetime.now().isoformat(timespec="seconds"),
                }
            else:
                # Blank clears the saved correction; it does not force future
                # source pulls to stay blank under the normal precedence rules.
                reconciled.pop(key, None)
    return ReviewedImport(reviewed, baseline, corrections_hash, reconciled,
                          proposal.confirmed_logon_changes, edited_fields,
                          proposal.summary, proposal.unresolved_count)


def persist_reviewed_import(store, plan: ReviewedImport) -> tuple[Path, Path]:
    """Commit an approved preview while the caller holds the shared-folder lock.

    Stage output files before promotion. Back up authoritative files and restore
    them if a replacement fails. This handles ordinary write errors; filesystem
    crashes or unavailable storage still require recovery from history backups.
    """
    if not store.has_baseline() or file_sha256(store.baseline_path) != plan.baseline.file_hash:
        raise RosterError("The baseline changed after preview. Select the reviewed workbook again.")
    if store._corrections_hash() != plan.corrections_hash:
        raise RosterError("Saved corrections changed after preview. Select the reviewed workbook again.")
    if file_sha256(plan.reviewed.source_path) != plan.reviewed.file_hash:
        raise RosterError("The reviewed workbook changed after preview. Save and select it again.")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup = store.history_dir / f"baseline_before_reviewed_import_{stamp}.xlsx"
    report = store.reports_dir / f"reviewed_roster_{stamp}{plan.reviewed.source_path.suffix.lower()}"
    corrections_backup = store.history_dir / f"corrections_before_reviewed_import_{stamp}.json"
    audit_backup = store.history_dir / f"audit_before_reviewed_import_{stamp}.csv"
    receipt_path = store.history_dir / f"reviewed_import_{stamp}.json"
    corrections_existed = store.overrides_path.exists()
    audit_existed = store.audit_path.exists()
    with TemporaryDirectory(prefix=".reviewed_import_", dir=store.root) as temporary:
        staging = Path(temporary)
        staged_baseline = write_baseline(staging / "baseline.xlsx", plan.reviewed.rows)
        staged_overrides = staging / "corrections.json"
        staged_overrides.write_text(json.dumps(plan.overrides, indent=2, ensure_ascii=False), encoding="utf-8")
        # Build the next audit bytes without touching the live log.
        entry = StringIO(newline="")
        writer = csv.writer(entry)
        if not audit_existed:
            writer.writerow(["Approved At", "Approved By", "Source File", "Previous Baseline Hash",
                             "New Baseline Hash", "Employee Count", "New Employees", "Missing Employees", "Review Items"])
        writer.writerow([datetime.now().isoformat(timespec="seconds"), getpass.getuser(),
                         plan.reviewed.source_path.name, plan.baseline.file_hash, file_sha256(staged_baseline),
                         len(plan.reviewed.rows), plan.summary["New employees"],
                         plan.summary["Missing from new pull"], plan.summary["Review items"]])
        staged_audit = staging / "audit.csv"
        staged_audit.write_bytes((store.audit_path.read_bytes() if audit_existed else b"") + entry.getvalue().encode("utf-8"))
        # Preserve the exact edited workbook as the record of this approval.
        shutil.copy2(plan.reviewed.source_path, report)
        if file_sha256(report) != plan.reviewed.file_hash:
            report.unlink()
            raise RosterError("The reviewed workbook changed during import. Save and select it again.")
        shutil.copy2(store.baseline_path, backup)
        if corrections_existed:
            shutil.copy2(store.overrides_path, corrections_backup)
        if audit_existed:
            shutil.copy2(store.audit_path, audit_backup)
        receipt = {
            "approved_by": getpass.getuser(), "approved_at": datetime.now().isoformat(timespec="seconds"),
            "reviewed_file": plan.reviewed.source_path.name, "reviewed_hash": plan.reviewed.file_hash,
            "baseline_before": plan.baseline.file_hash, "baseline_after": file_sha256(staged_baseline),
            "confirmed_logon_changes": plan.confirmed_links,
            "edited_fields": [change.__dict__ for change in plan.edits],
            "export_summary": plan.summary,
        }
        staged_receipt = staging / "receipt.json"
        staged_receipt.write_text(json.dumps(receipt, indent=2, ensure_ascii=False), encoding="utf-8")
        try:
            os.replace(staged_baseline, store.baseline_path)
            os.replace(staged_overrides, store.overrides_path)
            os.replace(staged_audit, store.audit_path)
            os.replace(staged_receipt, receipt_path)
        except OSError as exc:
            try:
                shutil.copy2(backup, store.baseline_path)
                for destination, saved, existed in [(store.overrides_path, corrections_backup, corrections_existed),
                                                    (store.audit_path, audit_backup, audit_existed)]:
                    if existed:
                        shutil.copy2(saved, destination)
                    else:
                        destination.unlink(missing_ok=True)
                report.unlink(missing_ok=True)
                receipt_path.unlink(missing_ok=True)
            except OSError as restore_error:
                raise RosterError(f"Import failed and recovery could not finish. Stop updating and restore from {backup.parent}: {restore_error}") from exc
            raise RosterError("Import failed; the previous baseline and corrections were restored. Close Excel files and try again.") from exc
    return report, backup
