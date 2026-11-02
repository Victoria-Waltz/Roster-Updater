"""Read saved capability corrections."""
from __future__ import annotations

from pathlib import Path
import json

from .models import RosterError
from .utils import clean, normalize_key


def load_overrides(path: Path) -> dict[str, dict[str, str]]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RosterError(f"Could not read saved capability corrections: {exc}") from exc
    if not isinstance(data, dict):
        raise RosterError("The saved capability corrections file is invalid.")
    return {
        normalize_key(key): {
            "capability": clean(value.get("capability")),
            "source": clean(value.get("source")) or "Manual",
            "notes": clean(value.get("notes")),
            "updated_by": clean(value.get("updated_by")),
            "updated_at": clean(value.get("updated_at")),
        }
        for key, value in data.items()
        if isinstance(value, dict) and normalize_key(key)
    }
