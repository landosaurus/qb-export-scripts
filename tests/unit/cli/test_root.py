from __future__ import annotations

from unittest.mock import MagicMock

from click.testing import CliRunner


def test_qb_help_lists_subcommands() -> None:
    from qb_cli.cli.root import qb

    runner = CliRunner()
    result = runner.invoke(qb, ["--help"], obj=MagicMock())
    assert result.exit_code == 0
    assert "export" in result.output
    assert "repl" in result.output
    assert "wizard" in result.output


def test_qb_no_subcommand_invokes_wizard(mocker) -> None:
    from qb_cli.cli.root import qb

    # Patch run_wizard so the test verifies wiring only — don't actually open a
    # questionary prompt (which fails on Windows CliRunner without a console).
    mock_run = mocker.patch("qb_cli.cli.wizard_cmd.run_wizard", return_value=0)

    runner = CliRunner()
    result = runner.invoke(qb, [], obj=MagicMock())
    assert result.exit_code == 0
    assert mock_run.call_count == 1
