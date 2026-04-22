from __future__ import annotations

import pytest

from qb_cli.repl.parser import parse_line


def test_empty_line_returns_empty_list() -> None:
    assert parse_line("") == []


def test_whitespace_only_returns_empty_list() -> None:
    assert parse_line("   ") == []


def test_comment_returns_empty_list() -> None:
    assert parse_line("# a comment") == []


def test_simple_argv_split() -> None:
    assert parse_line("export invoice --year 2025") == [
        "export",
        "invoice",
        "--year",
        "2025",
    ]


def test_quoted_argument_preserves_spaces() -> None:
    assert parse_line('export invoice --ref "some ref"') == [
        "export",
        "invoice",
        "--ref",
        "some ref",
    ]


def test_unclosed_quote_raises_value_error() -> None:
    with pytest.raises(ValueError):
        parse_line('export invoice --ref "unclosed')
