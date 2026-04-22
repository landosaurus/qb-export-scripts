from __future__ import annotations

import click

from qb_cli.context import Context


ENTITY_CHOICES = ["invoice", "sales_order", "purchase_order"]


@click.group(invoke_without_command=True)
@click.option(
    "--company-file",
    default=None,
    envvar="QB_CLI_COMPANY_FILE",
    help="Path to the QuickBooks company file (.QBW).",
)
@click.option("--config", "config_path", type=click.Path(), default=None)
@click.option(
    "--log-level",
    default="INFO",
    envvar="QB_CLI_LOG_LEVEL",
    show_default=True,
)
@click.option(
    "--json",
    "json_output",
    is_flag=True,
    default=False,
    help="Emit logs as single-line JSON.",
)
@click.pass_context
def qb(
    ctx: click.Context,
    company_file: str | None,
    config_path: str | None,
    log_level: str,
    json_output: bool,
) -> None:
    """Interactive REPL + non-interactive CLI for QuickBooks Desktop."""
    # Only build a Context from env when we don't already have one (tests pass
    # obj=fake_ctx via CliRunner.invoke). This lets unit tests bypass real config
    # loading while production use still wires a real Context.
    if ctx.obj is None:
        ctx.obj = Context.from_env(
            company_file=company_file,
            config_path=config_path,
            log_level=log_level,
            json_output=json_output,
        )
    if ctx.invoked_subcommand is None:
        ctx.invoke(repl)


# Sub-commands registered below (imports kept at module scope-bottom to avoid cycles).
from qb_cli.cli.export_cmd import export  # noqa: E402
from qb_cli.cli.import_cmd import import_cmd  # noqa: E402
from qb_cli.cli.query_cmd import query  # noqa: E402
from qb_cli.cli.repl_cmd import repl  # noqa: E402


qb.add_command(export)
qb.add_command(import_cmd)
qb.add_command(query)
qb.add_command(repl)
