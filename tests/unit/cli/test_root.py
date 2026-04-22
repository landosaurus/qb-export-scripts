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


def test_qb_no_subcommand_invokes_repl_stub() -> None:
    from qb_cli.cli.root import qb

    runner = CliRunner()
    result = runner.invoke(qb, [], obj=MagicMock())
    # repl stub exits 1 with the "not yet implemented" message on stderr
    assert result.exit_code == 1
    combined = result.output + (result.stderr if result.stderr_bytes else "")
    assert "REPL not yet implemented" in combined
