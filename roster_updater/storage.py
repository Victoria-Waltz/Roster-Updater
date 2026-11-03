"""Shared baseline persistence, corrections, locking, and approval history."""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterator
import csv
import getpass
import json
import os
import shutil
import time

from .comparison import compare_rosters
from .corrections import load_overrides
from .excel import write_baseline
from .loading import load_roster
from .identity import remap_corrections
from .models import ComparisonResult, LoadedRoster, RosterError
from .reports import write_report
from .reviewed_import import ReviewedImport, prepare_reviewed_import, persist_reviewed_import
from .utils import clean, file_sha256, normalize_key


@contextmanager
def shared_lock(path: Path, purpose: str, timeout_seconds: int = 0) -> Iterator[None]:
    """Hold an exclusive lock file while a shared-folder write is in progress.

    Use with shared_lock(...): around operations that must not overlap.
    A zero timeout fails immediately if the file exists; positive timeouts
    retry every quarter second. The lock is removed when the protected block
    exits, including on an exception. A process crash may leave a stale file.
    This relies on exclusive file creation on the same filesystem; independent
    local copies in sync folders do not provide a distributed lock.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    while True:
        try:
            # Exclusive creation makes checking and acquiring the lock one
            # filesystem operation, avoiding a check-then-create race.
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
    # Release after the protected operation even if it raises an exception.
    finally:
        try:
            path.unlink()
        except FileNotFoundError:
            pass


class SharedRosterStore:
    """Own the paths and file operations for one shared roster folder.

    The baseline is the approved roster; corrections persist independently.
    History retains backups/source pulls, reports retains approved reports,
    and the audit CSV records approvals. Constructing a store does not write.
    """
    def __init__(self, root: str | Path):
        """Resolve the shared folder and derive the standard data paths.

        root is the user-selected data folder, not the application source folder.
        """
        self.root = Path(root).expanduser().resolve()
        self.baseline_path = self.root / "current_baseline.xlsx"
        self.overrides_path = self.root / "capability_corrections.json"
        self.history_dir = self.root / "history"
        self.reports_dir = self.root / "reports"
        self.lock_path = self.root / ".roster_update.lock"
        self.audit_path = self.root / "approval_history.csv"

    def prepare(self) -> None:
        """Create the shared root, history, and reports directories if needed."""
        self.root.mkdir(parents=True, exist_ok=True)
        self.history_dir.mkdir(parents=True, exist_ok=True)
        self.reports_dir.mkdir(parents=True, exist_ok=True)

    def has_baseline(self) -> bool:
        """Check whether the baseline path exists without loading or validating it."""
        return self.baseline_path.exists()

    def load_baseline(self) -> LoadedRoster:
        """Load the current approved workbook, or explain that setup is still needed."""
        if not self.has_baseline():
            raise RosterError("The shared folder has not been initialized with a baseline yet.")
        return load_roster(self.baseline_path, preferred_sheet="Current Baseline")

    def initialize(self, baseline_source: str | Path) -> LoadedRoster:
        """Import an approved source as the shared baseline and return the saved roster.

        Validate the source first. Under the shared lock, back up an existing
        baseline before replacing it. Existing capability corrections are retained.
        """
        self.prepare()
        source = load_roster(baseline_source, preferred_sheet="All_Users_Active_Info (1)")
        with shared_lock(self.lock_path, "Initialize baseline"):
            if self.baseline_path.exists():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                shutil.copy2(self.baseline_path, self.history_dir / f"baseline_before_reinitialize_{timestamp}.xlsx")
            write_baseline(self.baseline_path, source.rows)
        return self.load_baseline()

    def load_overrides(self) -> dict[str, dict[str, str]]:
        """Load capability corrections from this store's JSON file."""
        return load_overrides(self.overrides_path)

    def save_override(
        self,
        logon_name: str,
        capability: str,
        source: str,
        notes: str = "",
    ) -> None:
        """Persist one approved capability correction with its source and audit details.

        Require nonblank employee key and capability. Under the lock, reload all
        corrections before updating one entry so unrelated entries are preserved.
        This changes the corrections file, not the baseline workbook.
        """
        key = normalize_key(logon_name)
        if not key:
            raise RosterError("A Logon Name is required to save a capability correction.")
        if not clean(capability):
            raise RosterError("Enter a capability before saving the correction.")
        self.prepare()
        with shared_lock(self.lock_path, "Save capability correction", timeout_seconds=3):
            # Reload inside the lock so another writer's saved corrections
            # are not lost when this employee's entry is changed.
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
        """Remove a saved correction if present, leaving other corrections intact.

        Future comparisons fall back to monthly or baseline values for this key.
        """
        key = normalize_key(logon_name)
        self.prepare()
        with shared_lock(self.lock_path, "Remove capability correction", timeout_seconds=3):
            # Reload inside the lock so another writer's saved corrections
            # are not lost when this employee's entry is changed.
            overrides = self.load_overrides()
            if key in overrides:
                overrides.pop(key)
                temp_path = self.overrides_path.with_suffix(".json.tmp")
                temp_path.write_text(json.dumps(overrides, indent=2, ensure_ascii=False), encoding="utf-8")
                os.replace(temp_path, self.overrides_path)

    def compare(
        self, monthly_path: str | Path,
        confirmed_logon_changes: dict[str, str] | None = None,
    ) -> ComparisonResult:
        """Load the current baseline, monthly file, and corrections into a proposal.

        No approval or file writes occur; the UI can review or export the result.
        """
        baseline = self.load_baseline()
        monthly = load_roster(monthly_path)
        corrections_hash = self._corrections_hash()
        overrides = self.load_overrides()
        if corrections_hash != self._corrections_hash():
            raise RosterError("Capability corrections changed while loading. Run the comparison again.")
        result = compare_rosters(baseline, monthly, overrides, confirmed_logon_changes)
        result.corrections_hash = corrections_hash
        return result

    def _corrections_hash(self) -> str:
        """Fingerprint the corrections snapshot; absence is a valid empty state."""
        return file_sha256(self.overrides_path) if self.overrides_path.exists() else ""

    def confirm_logon_change(self, result: ComparisonResult, old_logon: str, new_logon: str) -> ComparisonResult:
        """Revalidate the files and return a linked proposal without writing data.

        The comparison validates that each link remains a unique qualifying pair.
        Confirmation is scoped to this baseline/pull, not a permanent identity rule.
        """
        if not self.has_baseline() or file_sha256(self.baseline_path) != result.baseline.file_hash:
            raise RosterError("The baseline changed. Run the comparison again before confirming.")
        if file_sha256(result.monthly_pull.source_path) != result.monthly_pull.file_hash:
            raise RosterError("The monthly file changed. Run the comparison again before confirming.")
        if self._corrections_hash() != result.corrections_hash:
            raise RosterError("Capability corrections changed. Run the comparison again before confirming.")
        links = dict(result.confirmed_logon_changes)
        links[normalize_key(new_logon)] = normalize_key(old_logon)
        updated = self.compare(result.monthly_pull.source_path, links)
        # Protect against file replacement during the reload as well.
        if (updated.baseline.file_hash != result.baseline.file_hash
                or updated.monthly_pull.file_hash != result.monthly_pull.file_hash
                or updated.corrections_hash != result.corrections_hash):
            raise RosterError("The files changed during confirmation. Run the comparison again.")
        return updated

    def prepare_reviewed_import(self, path: str | Path) -> ReviewedImport:
        """Preview an edited app workbook before the user confirms promotion."""
        return prepare_reviewed_import(self, path)

    def import_reviewed_roster(self, plan: ReviewedImport) -> tuple[Path, Path]:
        """Import exactly the reviewed preview with backups and correction reconciliation."""
        self.prepare()
        with shared_lock(self.lock_path, "Import reviewed roster"):
            return persist_reviewed_import(self, plan)

    def approve(self, result: ComparisonResult) -> tuple[Path, Path]:
        """Promote a comparison to the shared baseline and return report/backup paths.

        Under the shared lock, reject a result if its baseline hash is stale.
        Then copy the old baseline and source pull, save a report, replace the
        baseline, and append an audit row. This sequence is not a multi-file
        transaction: a later failure does not roll back earlier completed writes.
        The UI handles confirmation and unresolved-item warnings; this method
        does not reject a result solely because it has unresolved review items.
        """
        self.prepare()
        with shared_lock(self.lock_path, "Approve monthly baseline"):
            if not self.baseline_path.exists():
                raise RosterError("The shared baseline no longer exists. Run the comparison again.")
            # A user may have left the review window open while someone else
            # approved a newer baseline. Reject that stale proposal.
            current_hash = file_sha256(self.baseline_path)
            if current_hash != result.baseline.file_hash:
                raise RosterError(
                    "The shared baseline changed after this comparison was run. "
                    "Another user may have approved an update. Run the comparison again before approving."
                )

            migrated_overrides = None
            if result.confirmed_logon_changes:
                if file_sha256(result.monthly_pull.source_path) != result.monthly_pull.file_hash:
                    raise RosterError("The monthly file changed. Run the comparison again before approving.")
                if result.corrections_hash != self._corrections_hash():
                    raise RosterError("Capability corrections changed. Run the comparison again before approving.")
                migrated_overrides = remap_corrections(self.load_overrides(), result.confirmed_logon_changes)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = self.history_dir / f"baseline_{timestamp}.xlsx"
            source_copy = self.history_dir / f"source_pull_{timestamp}{result.monthly_pull.source_path.suffix.lower()}"
            report_path = self.reports_dir / f"roster_update_{timestamp}.xlsx"
            # Preserve the previous baseline and the input pull before promotion.
            # Each step can fail independently; this is not an all-or-nothing transaction.
            shutil.copy2(self.baseline_path, backup_path)
            shutil.copy2(result.monthly_pull.source_path, source_copy)
            write_report(report_path, result)
            # Keep the old corrections for recovery before migrating confirmed keys.
            if migrated_overrides is not None and self.overrides_path.exists():
                shutil.copy2(self.overrides_path, self.history_dir / f"corrections_before_logon_change_{timestamp}.json")
            write_baseline(self.baseline_path, result.updated_roster)
            if migrated_overrides is not None:
                temp_path = self.overrides_path.with_suffix(".json.tmp")
                temp_path.write_text(json.dumps(migrated_overrides, indent=2, ensure_ascii=False), encoding="utf-8")
                os.replace(temp_path, self.overrides_path)
                # Record the human-confirmed mapping alongside this month's backups.
                receipt = {
                    "confirmed_by": getpass.getuser(),
                    "approved_at": datetime.now().isoformat(timespec="seconds"),
                    "baseline_hash": result.baseline.file_hash,
                    "monthly_hash": result.monthly_pull.file_hash,
                    "new_to_old_logon_names": result.confirmed_logon_changes,
                }
                (self.history_dir / f"confirmed_logon_changes_{timestamp}.json").write_text(
                    json.dumps(receipt, indent=2, ensure_ascii=False), encoding="utf-8",
                )

            new_hash = file_sha256(self.baseline_path)
            # Append one audit entry per approval; write headings only for a new log.
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


