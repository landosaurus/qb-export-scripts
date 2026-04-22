from __future__ import annotations

from datetime import date, timedelta

import pytest

from qb_cli.utils.date_parse import parse_date, year_range


def test_parse_iso_date():
    assert parse_date("2025-05-01") == date(2025, 5, 1)


def test_parse_us_date():
    assert parse_date("5/1/2025") == date(2025, 5, 1)


def test_parse_today():
    assert parse_date("today") == date.today()


def test_parse_yesterday():
    assert parse_date("yesterday") == date.today() - timedelta(days=1)


def test_parse_garbage_raises():
    with pytest.raises(ValueError):
        parse_date("garbage")


def test_year_range_past_year():
    assert year_range(2024, today=date(2026, 4, 22)) == (
        date(2024, 1, 1),
        date(2024, 12, 31),
    )


def test_year_range_current_year():
    assert year_range(2026, today=date(2026, 4, 22)) == (
        date(2026, 1, 1),
        date(2026, 4, 22),
    )
