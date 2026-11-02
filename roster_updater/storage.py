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
from .models import ComparisonResult, LoadedRoster, RosterError
from .reports import write_report
from .utils import clean, file_sha256, normalize_key


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
