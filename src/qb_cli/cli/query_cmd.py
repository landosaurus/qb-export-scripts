from __future__ import annotations

import click

from qb_cli.context import Context
from qb_cli.ops.query_op import query as _query


@click.command("query")
@click.argument(
    "entity", type=click.Choice(["invoice", "sales_order", "purchase_order"])
)
@click.argument("refs", nargs=-1)
@click.pass_obj
def query(ctx: Context, entity: str, refs: tuple[str, ...]) -> None:
    """Query records and print a summary to stdout."""
    records = _query(
        ctx, entity, ref_numbers=list(refs) if refs else None
    )
    for r in records:
        ref = getattr(r, "ref_number", "-") or "-"
        txn_date = getattr(r, "txn_date", "-")
        total = getattr(r, "total_amount", "-")
        click.echo(f"{ref}\t{txn_date}\t{total}")
    click.echo(f"\n{len(records)} {entity} record(s)", err=True)
