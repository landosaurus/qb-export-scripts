from __future__ import annotations

from datetime import date
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from click.testing import CliRunner
from pytest_mock import MockerFixture

from qb_cli.io.format import Format


def _fake_ctx() -> MagicMock:
    return MagicMock(name="FakeContext")


def test_export_invoice_ref_csv_extension(
    mocker: MockerFixture, tmp_path: Path
) -> None:
    from qb_cli.cli.root import qb

    out_path = tmp_path / "X.csv"
    mock_result = MagicMock(count=5, output_path=out_path)
    mock_export = mocker.patch(
        "qb_cli.cli.export_cmd._export", return_value=mock_result
    )

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["export", "invoice", "--ref", "14396", "--out", str(out_path)],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 0, result.output
    mock_export.assert_called_once()
    kwargs = mock_export.call_args.kwargs
    assert kwargs["ref_numbers"] == ["14396"]
    assert kwargs["output_path"] == out_path
    # fmt is None -> op infers from extension
    assert kwargs["fmt"] is None
    assert "exported 5" in result.output


def test_export_invoice_year_flag_populates_date_range(
    mocker: MockerFixture, tmp_path: Path
) -> None:
    from qb_cli.cli.root import qb
    from qb_cli.utils.date_parse import year_range

    out_path = tmp_path / "X.json"
    mock_result = MagicMock(count=7, output_path=out_path)
    mocker.patch("qb_cli.cli.export_cmd._export", return_value=mock_result)
    mock_export = mocker.patch(
        "qb_cli.cli.export_cmd._export", return_value=mock_result
    )

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["export", "invoice", "--year", "2025", "--out", str(out_path)],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 0, result.output
    expected_from, expected_to = year_range(2025)
    kwargs = mock_export.call_args.kwargs
    assert kwargs["date_from"] == expected_from
    assert kwargs["date_to"] == expected_to


def test_export_format_override_forces_json_despite_csv_extension(
    mocker: MockerFixture, tmp_path: Path
) -> None:
    from qb_cli.cli.root import qb

    out_path = tmp_path / "X.csv"
    mock_result = MagicMock(count=1, output_path=out_path)
    mock_export = mocker.patch(
        "qb_cli.cli.export_cmd._export", return_value=mock_result
    )

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["export", "invoice", "--out", str(out_path), "--format", "json"],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 0, result.output
    kwargs = mock_export.call_args.kwargs
    assert kwargs["fmt"] is Format.JSON


def test_export_rejects_unknown_entity(tmp_path: Path) -> None:
    from qb_cli.cli.root import qb

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["export", "widget", "--out", str(tmp_path / "X.csv")],
        obj=_fake_ctx(),
    )

    assert result.exit_code != 0
    assert "widget" in result.output or "invalid" in result.output.lower()


def test_export_from_and_to_parse_dates(
    mocker: MockerFixture, tmp_path: Path
) -> None:
    from qb_cli.cli.root import qb

    out_path = tmp_path / "X.json"
    mock_result = MagicMock(count=0, output_path=out_path)
    mock_export = mocker.patch(
        "qb_cli.cli.export_cmd._export", return_value=mock_result
    )

    runner = CliRunner()
    result = runner.invoke(
        qb,
        [
            "export",
            "invoice",
            "--from",
            "2025-01-15",
            "--to",
            "2025-02-20",
            "--out",
            str(out_path),
        ],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 0, result.output
    kwargs = mock_export.call_args.kwargs
    assert kwargs["date_from"] == date(2025, 1, 15)
    assert kwargs["date_to"] == date(2025, 2, 20)
