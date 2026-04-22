from __future__ import annotations

import shlex


def parse_line(line: str) -> list[str]:
    """Split an input line into argv tokens with shell-like quoting.

    Empty lines (after strip) return [].
    Comment lines starting with '#' return [].
    Raises ValueError on unclosed quotes.
    """
    stripped = line.strip()
    if not stripped or stripped.startswith("#"):
        return []
    return shlex.split(stripped)
