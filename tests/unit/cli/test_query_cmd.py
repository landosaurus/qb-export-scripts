from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

from click.testing import CliRunner
from pytest_mock import MockerFixture


def _fake_ctx() -> MagicMock:
    return MagicMock(name="FakeContext")


def test_query_prints_one_line_per_record(mocker: MockerFixture) -> None:
    from qb_cli.cli.root import qb

    records = [
        SimpleNamespace(ref_number="14396", txn_date="2025-01-05", total_amount="100.00"),
        SimpleNamespace(ref_number="14397", txn_date="2025-01-06", total_amount="200.50"),
    ]
    mock_query = mocker.patch(
        "qb_cli.cli.query_cmd._query", return_value=records
    )

    runner = CliRunner()
    result = runner.invoke(qb, ["query", "invoice"], obj=_fake_ctx())

    assert result.exit_code == 0, result.output
    assert "14396" in result.output
    assert "14397" in result.output
    mock_query.assert_called_once()
    kwargs = mock_query.call_args.kwargs
    assert kwargs["ref_numbers"] is None


def test_query_passes_refs(mocker: MockerFixture) -> None:
    from qb_cli.cli.root import qb

    mock_query = mocker.patch("qb_cli.cli.query_cmd._query", return_value=[])

    runner = CliRunner()
    result = runner.invoke(
        qb, ["query", "invoice", "14396", "14397"], obj=_fake_ctx()
    )

    assert result.exit_code == 0, result.output
    kwargs = mock_query.call_args.kwargs
    assert kwargs["ref_numbers"] == ["14396", "14397"]


def test_query_rejects_unknown_entity() -> None:
    from qb_cli.cli.root import qb

    runner = CliRunner()
    result = runner.invoke(qb, ["query", "widget"], obj=_fake_ctx())

    assert result.exit_code != 0
