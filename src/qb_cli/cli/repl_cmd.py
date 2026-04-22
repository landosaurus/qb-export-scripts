from __future__ import annotations

import click

from qb_cli.context import Context
from qb_cli.repl.shell import run_repl


@click.command("repl")
@click.pass_obj
def repl(ctx: Context) -> None:
    """Launch the interactive REPL shell."""
    raise click.exceptions.Exit(run_repl(ctx))
