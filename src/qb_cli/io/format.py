from __future__ import annotations

from enum import Enum
from pathlib import Path


class Format(str, Enum):
    CSV = "csv"
    JSON = "json"


def detect_format(path: str | Path) -> Format:
    ext = Path(path).suffix.lower().lstrip(".")
    try:
        return Format(ext)
    except ValueError as e:
        raise ValueError(f"unsupported file format '{ext}' for {path}") from e
