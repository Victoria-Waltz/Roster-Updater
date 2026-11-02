"""Text normalization and file hashing."""
from __future__ import annotations

from pathlib import Path
from typing import Any
import hashlib
import re

from .constants import HEADER_ALIASES


def clean(value: Any) -> str:
    """Convert a spreadsheet value to trimmed text.

    Blank cells become empty strings, and whole-number floats lose their .0
    suffix so Excel formatting does not introduce false differences.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_key(value: Any) -> str:
    """Return trimmed, case-insensitive text for matching and comparisons.

    Use this for dictionary keys; keep clean() values for readable output.
    """
    return clean(value).casefold()


def normalize_header(value: Any) -> str:
    """Map a heading to its canonical column name using HEADER_ALIASES.

    Punctuation, spaces, and case are ignored for alias lookup. Unknown
    headings retain their trimmed spelling so validation can report them.
    """
    compact = re.sub(r"[^a-z0-9]", "", clean(value).casefold())
    return HEADER_ALIASES.get(compact, clean(value))


def file_sha256(path: Path) -> str:
    """Hash file bytes in 1 MiB chunks without loading the whole file into memory.

    The returned fingerprint lets approval detect that a baseline has changed.
    """
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()
