from __future__ import annotations

import os
from pathlib import Path

import click
from prompt_toolkit import PromptSession
from prompt_toolkit.history import FileHistory
from prompt_toolkit.input import Input
from prompt_toolkit.output import Output

from qb_cli.context import Context
from qb_cli.repl.completers import QBCompleter
from qb_cli.repl.parser import parse_line


EXIT_TOKENS: set[str] = {"exit", "quit"}


def _history_path() -> Path:
    """Platform-appropriate history file."""
    if os.name == "nt":
        base = os.environ.get("APPDATA")
        root = Path(base) if base else Path.home() / "AppData" / "Roaming"
    else:
        xdg = os.environ.get("XDG_STATE_HOME")
        root = Path(xdg) if xdg else Path.home() / ".local" / "state"
    d = root / "qb_cli"
    d.mkdir(parents=True, exist_ok=True)
    return d / "history"


def _build_prompt(ctx: Context) -> str:
    suffix_bits: list[str] = []
    if getattr(ctx, "dry_run", False):
        suffix_bits.append("dry")
    suffix = f"[{','.join(suffix_bits)}] " if suffix_bits else ""
    return f"qb> {suffix}"


def run_repl(
    ctx: Context,
    *,
    input: Input | None = None,
    output: Output | None = None,
) -> int:
    """Run the interactive REPL. Returns the intended process exit code.

    The `input`/`output` parameters are for testability; when None, PromptSession picks stdin/stdout.
    """
    session: PromptSession[str] = PromptSession(
        message=_build_prompt(ctx),
        history=FileHistory(str(_history_path())),
        completer=QBCompleter(),
        input=input,
        output=output,
    )
    # Import here to avoid circular dep (cli.root imports repl_cmd which imports us).
    from qb_cli.cli.root import qb as _qb_group

    while True:
        try:
            line = session.prompt()
        except (EOFError, KeyboardInterrupt):
            return 0

        try:
            argv = parse_line(line)
        except ValueError as e:
            click.echo(f"parse error: {e}", err=True)
            continue

        if not argv:
            continue

        if argv[0] in EXIT_TOKENS:
            return 0

        # Dispatch through the click group, reusing the same Context object.
        try:
            _qb_group.main(args=argv, standalone_mode=False, obj=ctx)
        except click.exceptions.Exit as e:
            # Non-zero exit from a sub-command shouldn't kill the REPL.
            if e.exit_code != 0:
                click.echo(f"command exited {e.exit_code}", err=True)
        except click.ClickException as e:
            e.show()
        except SystemExit as e:
            # Click sometimes raises SystemExit despite standalone_mode=False (e.g. --help).
            if e.code not in (0, None):
                click.echo(f"command exited {e.code}", err=True)
        except Exception as e:  # noqa: BLE001 — REPL must not die on user typos
            click.echo(f"error: {e}", err=True)
