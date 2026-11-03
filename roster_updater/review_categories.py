"""Shared review labels and colors for the desktop list and Excel report."""

CATEGORY_LABELS = {
    "changed": "Information changed",
    "missing": "Missing from new pull",
    "new": "New employee",
    "manager": "Manager issue",
    "capability": "Capability review",
}

CATEGORY_COLORS = {
    "changed": "D9EAF7",
    "missing": "F4CCCC",
    "new": "E2F0D9",
    "manager": "FCE4D6",
    "capability": "FFF2CC",
}


def category_for_issue(issue: str) -> str:
    """Classify a finding consistently regardless of capitalization.

    Missing employees take priority over other words in an issue description.
    Capability suggestions belong to capability review even when a manager
    supplied the suggestion. Unknown findings use the information-change color.
    """
    text = issue.casefold()
    if "missing from" in text:
        return "missing"
    if "capability" in text:
        return "capability"
    if "manager" in text:
        return "manager"
    if "new employee" in text:
        return "new"
    return "changed"
