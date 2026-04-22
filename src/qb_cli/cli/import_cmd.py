from __future__ import annotations

from pathlib import Path
from typing import cast

import click

from qb_cli.context import Context
from qb_cli.io.format import Format
from qb_cli.ops.import_op import OnDuplicate, import_ as _import


@click.command("import")
@click.argument(
    "entity", type=click.Choice(["invoice", "sales_order", "purchase_order"])
)
@click.option(
    "--file",
    "input_path",
    type=click.Path(exists=True, dir_okay=False),
    required=True,
    help="Input file (.csv or .json).",
)
@click.option(
    "--format",
    "fmt",
    type=click.Choice(["csv", "json"]),
    default=None,
)
@click.option(
    "--dry-run",
    is_flag=True,
    default=False,
    help="Run pre-flight and dedupe but do not write to QuickBooks.",
)
@click.option(
    "--on-duplicate",
    type=click.Choice(["error", "skip", "update"]),
    default="error",
    show_default=True,
)
@click.pass_obj
def import_cmd(
    ctx: Context,
    entity: str,
    input_path: str,
    fmt: str | None,
    dry_run: bool,
    on_duplicate: str,
) -> None:
    """Import records from a CSV or JSON file into QuickBooks."""
    format_enum = Format(fmt) if fmt else None
    # click.Choice restricts values to the OnDuplicate literal set, so the cast
    # is sound at runtime and preserves the Literal type at the op boundary.
    on_duplicate_literal = cast(OnDuplicate, on_duplicate)
    result = _import(
        ctx,
        entity,
        input_path=Path(input_path),
        fmt=format_enum,
        dry_run=dry_run,
        on_duplicate=on_duplicate_literal,
    )
    if result.missing_refs:
        click.echo(
            f"missing references — aborted: {', '.join(result.missing_refs)}",
            err=True,
        )
        raise click.exceptions.Exit(2)
    if result.duplicate_refs and on_duplicate == "error":
        click.echo(
            f"duplicates found — aborted: {', '.join(result.duplicate_refs)}",
            err=True,
        )
        raise click.exceptions.Exit(3)
    action = "would import" if dry_run else "imported"
    click.echo(
        f"{action} {result.written} "
        f"(attempted {result.attempted}, skipped {result.skipped}, "
        f"failed {result.failed})"
    )
    if result.errors:
        for ref, msg in result.errors:
            click.echo(f"  [{ref}] {msg}", err=True)
        raise click.exceptions.Exit(1)
