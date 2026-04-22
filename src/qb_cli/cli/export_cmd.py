from __future__ import annotations

from datetime import date
from pathlib import Path

import click

from qb_cli.context import Context
from qb_cli.io.format import Format
from qb_cli.ops.export_op import export as _export
from qb_cli.utils.date_parse import parse_date, year_range


@click.command("export")
@click.argument(
    "entity", type=click.Choice(["invoice", "sales_order", "purchase_order"])
)
@click.option(
    "--ref",
    "ref_numbers",
    multiple=True,
    help="Fetch one or more specific RefNumbers.",
)
@click.option(
    "--year",
    type=int,
    help="Shortcut: fetch all records for YEAR (Jan 1 through Dec 31 or today).",
)
@click.option(
    "--from",
    "date_from",
    type=str,
    help="Start date (YYYY-MM-DD, MM/DD/YYYY, today, yesterday).",
)
@click.option("--to", "date_to", type=str, help="End date.")
@click.option(
    "--out",
    "output_path",
    type=click.Path(dir_okay=False, writable=True),
    required=True,
    help="Output file path. Extension (.csv or .json) implies format.",
)
@click.option(
    "--format",
    "fmt",
    type=click.Choice(["csv", "json"]),
    default=None,
    help="Override format (defaults to extension).",
)
@click.pass_obj
def export(
    ctx: Context,
    entity: str,
    ref_numbers: tuple[str, ...],
    year: int | None,
    date_from: str | None,
    date_to: str | None,
    output_path: str,
    fmt: str | None,
) -> None:
    """Export records from QuickBooks to CSV or JSON."""
    df: date | None = None
    dt: date | None = None
    if year is not None:
        df, dt = year_range(year)
    if date_from is not None:
        df = parse_date(date_from)
    if date_to is not None:
        dt = parse_date(date_to)

    format_enum = Format(fmt) if fmt else None
    result = _export(
        ctx,
        entity,
        ref_numbers=list(ref_numbers) if ref_numbers else None,
        date_from=df,
        date_to=dt,
        output_path=Path(output_path),
        fmt=format_enum,
    )
    click.echo(f"exported {result.count} {entity} records to {result.output_path}")
