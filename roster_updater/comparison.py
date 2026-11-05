"""Compare employee records and identify review items."""
from __future__ import annotations

from typing import Any, Iterable

from .constants import (
    BASELINE_HEADERS,
    CAPABILITY_HEADER,
    COMPARE_FIELDS,
    ERP_HEADER,
    KEY_HEADER,
)
from .identity import find_logon_candidates, remap_corrections, validate_logon_links
from .models import ComparisonResult, FieldChange, LoadedRoster, ReviewItem
from .utils import clean, normalize_comparison_value, normalize_key


def _row_map(rows: Iterable[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """Index records by normalized Logon Name for fast employee matching.

    Input should already be validated; duplicate keys would overwrite earlier
    records here. Blank keys are excluded.
    """
    return {normalize_key(row[KEY_HEADER]): row for row in rows if normalize_key(row[KEY_HEADER])}


def _build_people_indexes(
    rows: Iterable[dict[str, Any]],
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]], set[str]]:
    """Build email and unambiguous-name lookups for manager matching.

    Return email lookup, name lookup, and the set of repeated display names.
    Repeated names are removed from the name lookup to avoid guessing which
    person is intended. Repeated emails currently keep the last record.
    """
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
    # A display name is not unique enough when two employees share it.
    # Keep email matching available, but disable ambiguous name fallback.
    for name in duplicate_names:
        by_name.pop(name, None)
    return by_email, by_name, duplicate_names


def _resolve_manager(
    employee: dict[str, Any],
    by_email: dict[str, dict[str, Any]],
    by_name: dict[str, dict[str, Any]],
) -> dict[str, Any] | None:
    """Find a manager by email first, then by unambiguous display name.

    Return None when neither lookup finds a match in the supplied roster.
    """
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
    """Return the direct manager and that manager's manager, when available.

    This performs at most two lookups; it does not traverse the entire hierarchy.
    """
    manager = _resolve_manager(employee, by_email, by_name)
    managers_manager = _resolve_manager(manager, by_email, by_name) if manager else None
    return manager, managers_manager


def compare_rosters(
    baseline: LoadedRoster,
    monthly_pull: LoadedRoster,
    overrides: dict[str, dict[str, str]] | None = None,
    confirmed_logon_changes: dict[str, str] | None = None,
) -> ComparisonResult:
    """Build an updated roster and review findings without writing any files.

    Args:
        baseline: Previously approved, validated employee records.
        monthly_pull: New validated active-user records.
        overrides: Saved corrections keyed by normalized logon name, or None.
        confirmed_logon_changes: Explicit new-key -> old-key links for these files.
            Candidates alone never transfer baseline values or corrections.

    Returns:
        A ComparisonResult with proposed rows, review items, field changes,
        capability provenance, and summary counts. Input rows are not modified.

    Capability precedence is saved correction, nonblank monthly value, then
    nonblank baseline value. Manager capabilities are suggestions only.
    The proposed roster contains only people present in the monthly pull;
    missing baseline employees receive review items as possible departures.
    Field changes describe baseline versus raw monthly values, not overrides.
    """
    overrides = overrides or {}
    baseline_map = _row_map(baseline.rows)
    monthly_map = _row_map(monthly_pull.rows)
    candidates = find_logon_candidates(baseline.rows, monthly_pull.rows)
    links = validate_logon_links(candidates, confirmed_logon_changes)
    overrides = remap_corrections(overrides, links)

    updated: list[dict[str, Any]] = []
    capability_sources: dict[str, str] = {}
    changes: list[FieldChange] = []
    review: list[ReviewItem] = []
    changed_people: set[str] = set()

    # Pass 1: build proposed records and compare the raw monthly fields.
    # Create new dictionaries so the original input rows stay available for audit.
    for source_row in monthly_pull.rows:
        key = normalize_key(source_row[KEY_HEADER])
        prior = baseline_map.get(links.get(key, key))
        final_row = {header: clean(source_row.get(header)) for header in BASELINE_HEADERS}
        override = overrides.get(key, {})
        override_capability = clean(override.get("capability"))
        pull_capability = clean(source_row.get(CAPABILITY_HEADER))
        baseline_capability = clean(prior.get(CAPABILITY_HEADER)) if prior else ""

        # Approved corrections win across months. A blank incoming capability
        # must not erase a previously known baseline value.
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

        # ERP Category belongs to the baseline; the monthly pull does not
        # replace it. New employees begin with an empty ERP Category.
        final_row[ERP_HEADER] = clean(prior.get(ERP_HEADER)) if prior else ""
        updated.append(final_row)

        if prior:
            person_changes: list[str] = []
            # Logon Name is normally the key, but a confirmed rename is itself
            # a field change and must remain visible in the exported audit trail.
            fields = ([KEY_HEADER] if key in links else []) + COMPARE_FIELDS
            for field_name in fields:
                old_value = clean(prior.get(field_name))
                raw_new_value = clean(source_row.get(field_name))
                # A missing monthly capability is handled by carry-forward,
                # so do not report it as an intentional deletion.
                if field_name == CAPABILITY_HEADER and not normalize_comparison_value(raw_new_value):
                    continue
                if normalize_comparison_value(old_value) != normalize_comparison_value(raw_new_value):
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
    # Manager lookups use the completed proposed roster, allowing a manager
    # whose capability came from a saved correction to provide a suggestion.
    by_email, by_name, _ = _build_people_indexes(updated)

    capability_suggestions = 0
    missing_capabilities = 0
    manager_issue_people: set[str] = set()

    # Pass 2: inspect manager relationships and missing capabilities.
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
            # Add advice to the review list only. The employee remains blank
            # until a person explicitly saves/applies the correction in the UI.
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

    # Absence from the active pull is a review signal, not proof of departure.
    # These people are not included in the proposed active roster.
    departed_keys = [key for key in baseline_map if key not in monthly_map and key not in links.values()]
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

    pending_candidates = [item for item in candidates if normalize_key(item.new_logon) not in links]
    for candidate in pending_candidates:
        review.append(ReviewItem(
            logon_name=candidate.new_logon,
            display_name=candidate.display_name,
            issue="Possible Logon Name Change",
            current_value=candidate.old_logon,
            suggested_value=candidate.new_logon,
            changed_fields=KEY_HEADER,
            notes="Matching evidence: " + "; ".join(candidate.evidence_details) + ". " + (
                "Ambiguous: multiple qualifying records. Do not link; verify the source files."
                if candidate.ambiguous else
                "Select this finding and use Confirm logon change after verifying the employee."
            ),
        ))
    for new, old in links.items():
        review.append(ReviewItem(
            logon_name=clean(monthly_map[new][KEY_HEADER]),
            display_name=clean(monthly_map[new].get("Display Name")),
            issue="Confirmed Logon Name Change",
            current_value=clean(baseline_map[old][KEY_HEADER]),
            suggested_value=clean(monthly_map[new][KEY_HEADER]),
            changed_fields=KEY_HEADER,
            notes="Explicitly confirmed for this comparison. ERP Category and capability precedence now use the prior record. "
                  "Saved corrections move only when this month is approved in the app.",
        ))

    # Enrich findings with hierarchy context useful to the person reviewing them.
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

    new_keys = set(monthly_map) - set(baseline_map) - set(links)
    # Sets count unique people; changes/review lists count individual findings.
    # One employee can therefore contribute several review items.
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
        "Possible logon changes": len(pending_candidates),
        "Confirmed logon changes": len(links),
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
        logon_candidates=pending_candidates,
        confirmed_logon_changes=links,
    )
