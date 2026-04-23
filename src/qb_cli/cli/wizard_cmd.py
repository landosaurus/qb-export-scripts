from __future__ import annotations

import click

from qb_cli.context import Context
from qb_cli.repl.wizard import run_wizard


@click.command("wizard")
@click.pass_obj
def wizard(ctx: Context) -> None:
    """Launch the interactive menu-driven wizard."""
    raise click.exceptions.Exit(run_wizard(ctx))
