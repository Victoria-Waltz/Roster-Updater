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
