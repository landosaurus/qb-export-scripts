from __future__ import annotations

import click

from qb_cli.context import Context


@click.command("repl")
@click.pass_obj
def repl(ctx: Context) -> None:  # noqa: ARG001 — ctx is used when REPL lands in Task 11
    """Launch the interactive REPL shell."""
    click.echo("REPL not yet implemented (Task 11).", err=True)
    raise click.exceptions.Exit(1)
