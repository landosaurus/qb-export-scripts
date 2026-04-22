from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Tuple


_ISO_FORMAT = "%Y-%m-%d"
_US_FORMAT = "%m/%d/%Y"


def parse_date(s: str) -> date:
    """Accept YYYY-MM-DD, MM/DD/YYYY, 'today', 'yesterday'. Raise ValueError otherwise."""
    token = s.strip()
    lower = token.lower()
    if lower == "today":
        return date.today()
    if lower == "yesterday":
        return date.today() - timedelta(days=1)
    for fmt in (_ISO_FORMAT, _US_FORMAT):
        try:
            return datetime.strptime(token, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"could not parse date: {s!r}")


def year_range(year: int, today: date | None = None) -> Tuple[date, date]:
    """(YYYY-01-01, today) if year == current_year else (YYYY-01-01, YYYY-12-31).

    `today` is an injectable override for testability; defaults to real today.
    """
    effective_today = today if today is not None else date.today()
    start = date(year, 1, 1)
    if year == effective_today.year:
        return start, effective_today
    return start, date(year, 12, 31)
