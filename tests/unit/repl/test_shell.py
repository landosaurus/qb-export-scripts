from __future__ import annotations

from unittest.mock import MagicMock

import pytest
from prompt_toolkit.input import create_pipe_input
from prompt_toolkit.output import DummyOutput
from pytest_mock import MockerFixture

from qb_cli.repl.shell import run_repl


def _run(lines: list[str], ctx: object) -> int:
    """Drive run_repl with a pre-baked list of input lines."""
    with create_pipe_input() as inp:
        for line in lines:
            inp.send_text(line + "\r")
        # Close sends an EOF which ends the session cleanly (EOFError).
        inp.close()
        return run_repl(ctx, input=inp, output=DummyOutput())


def test_exit_cleanly() -> None:
    ctx = MagicMock()
    ctx.dry_run = False
    assert _run(["exit"], ctx) == 0


def test_quit_cleanly() -> None:
    ctx = MagicMock()
    ctx.dry_run = False
    assert _run(["quit"], ctx) == 0


def test_ctrl_d_cleanly() -> None:
    ctx = MagicMock()
    ctx.dry_run = False
    # No input lines — the immediate EOF should terminate cleanly.
    assert _run([], ctx) == 0


def test_parse_error_does_not_exit(capsys: pytest.CaptureFixture[str]) -> None:
    ctx = MagicMock()
    ctx.dry_run = False
    # Unclosed quote on line 1, then exit on line 2
    rc = _run(['export invoice --ref "unclosed', "exit"], ctx)
    assert rc == 0
    err = capsys.readouterr().err
    assert "parse error" in err


def test_empty_line_loop() -> None:
    ctx = MagicMock()
    ctx.dry_run = False
    rc = _run(["", "   ", "# comment", "exit"], ctx)
    assert rc == 0


def test_dispatch_to_cli(mocker: MockerFixture) -> None:
    ctx = MagicMock()
    ctx.dry_run = False
    # Patch the root qb group's main so we don't exercise the whole CLI stack
    fake_main = mocker.patch("qb_cli.cli.root.qb.main")
    _run(["help", "exit"], ctx)
    assert fake_main.call_count == 1
    _, kwargs = fake_main.call_args_list[0]
    assert kwargs["args"] == ["help"]
    assert kwargs["standalone_mode"] is False
    assert kwargs["obj"] is ctx
