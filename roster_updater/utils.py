"""Text normalization and file hashing."""
from __future__ import annotations

from pathlib import Path
from typing import Any
import hashlib
import re

from .constants import HEADER_ALIASES


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_key(value: Any) -> str:
    return clean(value).casefold()


def normalize_header(value: Any) -> str:
    compact = re.sub(r"[^a-z0-9]", "", clean(value).casefold())
    return HEADER_ALIASES.get(compact, clean(value))


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()
