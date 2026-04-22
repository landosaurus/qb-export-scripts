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


def test_qb_no_subcommand_invokes_repl() -> None:
    from qb_cli.cli.root import qb

    runner = CliRunner()
    # With no input piped, the REPL sees immediate EOF and exits cleanly.
    result = runner.invoke(qb, [], obj=MagicMock(), input="")
    assert result.exit_code == 0
