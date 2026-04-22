from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

from click.testing import CliRunner
from pytest_mock import MockerFixture

from qb_cli.ops.import_op import ImportResult


def _fake_ctx() -> MagicMock:
    return MagicMock(name="FakeContext")


def _touch(path: Path, content: str = "[]") -> Path:
    path.write_text(content)
    return path


def test_import_happy_path(mocker: MockerFixture, tmp_path: Path) -> None:
    from qb_cli.cli.root import qb

    in_path = _touch(tmp_path / "in.json")
    result_obj = ImportResult(entity="invoice", attempted=5, written=5)
    mock_import = mocker.patch(
        "qb_cli.cli.import_cmd._import", return_value=result_obj
    )

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["import", "invoice", "--file", str(in_path)],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 0, result.output
    mock_import.assert_called_once()
    kwargs = mock_import.call_args.kwargs
    assert kwargs["dry_run"] is False
    assert kwargs["on_duplicate"] == "error"
    assert "imported 5" in result.output


def test_import_missing_refs_exits_2(
    mocker: MockerFixture, tmp_path: Path
) -> None:
    from qb_cli.cli.root import qb

    in_path = _touch(tmp_path / "in.json")
    result_obj = ImportResult(
        entity="invoice",
        attempted=0,
        written=0,
        missing_refs=["customer:ACME"],
    )
    mocker.patch("qb_cli.cli.import_cmd._import", return_value=result_obj)

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["import", "invoice", "--file", str(in_path)],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 2
    assert "missing" in result.stderr.lower()
    assert "customer:ACME" in result.stderr


def test_import_duplicates_with_default_policy_exits_3(
    mocker: MockerFixture, tmp_path: Path
) -> None:
    from qb_cli.cli.root import qb

    in_path = _touch(tmp_path / "in.json")
    result_obj = ImportResult(
        entity="invoice",
        attempted=0,
        written=0,
        duplicate_refs=["14396"],
    )
    mocker.patch("qb_cli.cli.import_cmd._import", return_value=result_obj)

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["import", "invoice", "--file", str(in_path)],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 3
    assert "duplicate" in result.stderr.lower()
    assert "14396" in result.stderr


def test_import_dry_run_flag_is_forwarded(
    mocker: MockerFixture, tmp_path: Path
) -> None:
    from qb_cli.cli.root import qb

    in_path = _touch(tmp_path / "in.json")
    result_obj = ImportResult(entity="invoice", attempted=2, written=0)
    mock_import = mocker.patch(
        "qb_cli.cli.import_cmd._import", return_value=result_obj
    )

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["import", "invoice", "--file", str(in_path), "--dry-run"],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 0, result.output
    kwargs = mock_import.call_args.kwargs
    assert kwargs["dry_run"] is True
    assert "would import" in result.output


def test_import_errors_exit_1(mocker: MockerFixture, tmp_path: Path) -> None:
    from qb_cli.cli.root import qb

    in_path = _touch(tmp_path / "in.json")
    result_obj = ImportResult(
        entity="invoice",
        attempted=2,
        written=1,
        failed=1,
        errors=[("14396", "boom")],
    )
    mocker.patch("qb_cli.cli.import_cmd._import", return_value=result_obj)

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["import", "invoice", "--file", str(in_path)],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 1
    assert "14396" in result.stderr
    assert "boom" in result.stderr
