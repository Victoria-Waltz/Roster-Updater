"""Per-user desktop preferences."""
from __future__ import annotations

from pathlib import Path
import json
import os


def user_config_path() -> Path:
    """Locate per-user config.json under LOCALAPPDATA on Windows or ~/.config.

    These preferences are separate from the shared employee data folder.
    """
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    else:
        base = Path.home() / ".config"
    return base / "TAITRosterUpdater" / "config.json"


class ConfigStore:
    """Read and save local preferences such as the last chosen shared folder."""
    def __init__(self) -> None:
        """Remember the platform-specific settings path without creating a file."""
        self.path = user_config_path()

    def load(self) -> dict[str, str]:
        """Return saved settings, or an empty dict for missing, unreadable, or invalid JSON."""
        if not self.path.exists():
            return {}
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}
        return data if isinstance(data, dict) else {}

    def save(self, data: dict[str, str]) -> None:
        """Create the preferences directory and replace settings via a temporary file."""
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.path.with_suffix(".tmp")
        temp.write_text(json.dumps(data, indent=2), encoding="utf-8")
        os.replace(temp, self.path)
