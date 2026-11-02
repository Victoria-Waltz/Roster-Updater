"""Shared roster data models and user-facing errors."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any


class RosterError(Exception):
    """A user-facing roster processing error."""


@dataclass
class LoadedRoster:
    """A validated input roster plus the metadata needed to trace its source.

    rows use canonical column names; headers describe the source headings.
    file_hash records the loaded file's bytes for later stale-baseline checks.
    """
    rows: list[dict[str, Any]]
    source_path: Path
    sheet_name: str
    headers: list[str]
    file_hash: str


@dataclass
class ReviewItem:
    """One finding displayed in the review list and exported to Excel.

    One employee may have several findings. suggested_value is advice only;
    it does not change the proposed roster until a correction is saved.
    """
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
        """Return fields in the exact order expected by the report's Review headers."""
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
    """One changed field, preserving readable old and new source values."""
    logon_name: str
    display_name: str
    field_name: str
    old_value: str
    new_value: str

    def as_row(self) -> list[str]:
        """Return the five values expected by the report's Changes sheet."""
        return [
            self.logon_name,
            self.display_name,
            self.field_name,
            self.old_value,
            self.new_value,
        ]


@dataclass
class ComparisonResult:
    """In-memory proposal shared by comparison, the UI, reporting, and approval.

    Creating this object does not approve a month. capability_sources maps
    normalized employee keys to the origin of their selected capability.
    """
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
        """Count capability and manager findings that trigger the approval warning.

        This counts review items, not unique employees, and excludes informational
        new-employee, changed-field, and missing-from-pull findings.
        """
        unresolved = {
            "Missing capability",
            "Capability suggestion",
            "Missing manager information",
            "Manager not in active pull",
        }
        return sum(item.issue in unresolved for item in self.review_items)
