"""Suggest conservative employee-key changes without automatically linking records."""
from collections import Counter, defaultdict

from .constants import KEY_HEADER
from .models import LogonNameCandidate, RosterError
from .utils import clean, normalize_key

EVIDENCE_FIELDS = ("Title", "Department", "Manager", "Manager Email Address")


def find_logon_candidates(baseline_rows, monthly_rows) -> list[LogonNameCandidate]:
    """Look only at otherwise unmatched records with the same nonblank display name.

    Require at least two nonblank matching supporting fields. This is a review
    heuristic, not proof of identity. Any old/new record involved in multiple
    qualifying pairs is ambiguous and cannot be linked through confirmation.
    """
    baseline = {normalize_key(row[KEY_HEADER]): row for row in baseline_rows}
    monthly = {normalize_key(row[KEY_HEADER]): row for row in monthly_rows}
    old_by_name = defaultdict(list)
    for key in baseline.keys() - monthly.keys():
        name = normalize_key(baseline[key].get("Display Name"))
        if name:
            old_by_name[name].append(baseline[key])
    pairs = []
    for key in sorted(monthly.keys() - baseline.keys()):
        new = monthly[key]
        for old in old_by_name[normalize_key(new.get("Display Name"))]:
            evidence = tuple(field for field in EVIDENCE_FIELDS
                             if normalize_key(old.get(field))
                             and normalize_key(old.get(field)) == normalize_key(new.get(field)))
            if len(evidence) >= 2:
                pairs.append((old, new, evidence))
    old_counts = Counter(normalize_key(old[KEY_HEADER]) for old, _, _ in pairs)
    new_counts = Counter(normalize_key(new[KEY_HEADER]) for _, new, _ in pairs)
    return sorted([
        LogonNameCandidate(
            clean(old[KEY_HEADER]), clean(new[KEY_HEADER]), clean(new.get("Display Name")),
            ("Display Name",) + evidence,
            old_counts[normalize_key(old[KEY_HEADER])] > 1 or new_counts[normalize_key(new[KEY_HEADER])] > 1,
            tuple(f"{field}: {clean(new.get(field))}" for field in ("Display Name",) + evidence),
        )
        for old, new, evidence in pairs
    ], key=lambda item: (normalize_key(item.new_logon), normalize_key(item.old_logon)))


def validate_logon_links(candidates, confirmed_links) -> dict[str, str]:
    """Validate an explicitly confirmed new-key -> old-key mapping.

    Reject weak, ambiguous, stale, blank, or many-to-one requests rather than
    transferring an unrelated person's data. Revalidate on every comparison.
    """
    valid = {(normalize_key(item.new_logon), normalize_key(item.old_logon))
             for item in candidates if not item.ambiguous}
    links = {}
    for new, old in (confirmed_links or {}).items():
        new, old = normalize_key(new), normalize_key(old)
        if (new, old) not in valid or new in links or old in links.values():
            raise RosterError("This Logon Name link is ambiguous or no longer matches the files. Run a fresh comparison.")
        links[new] = old
    return links


def remap_corrections(overrides, links):
    """Copy corrections to confirmed new keys; retain an existing new-key correction.

    Do not modify the caller's dictionary. Storage persists this mapping only
    when the month is approved, so a preview can be cleared without side effects.
    """
    updated = dict(overrides)
    for new, old in links.items():
        if old in updated:
            updated.setdefault(new, updated[old])
            del updated[old]
    return updated
