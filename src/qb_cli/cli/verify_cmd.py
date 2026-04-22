from __future__ import annotations

import click

from qb_cli.context import Context
from qb_cli.ops.verify_op import verify_entities as _verify


@click.command("verify")
@click.option("--customer", "customers", multiple=True)
@click.option("--vendor", "vendors", multiple=True)
@click.option("--item", "items", multiple=True)
@click.option("--terms", "terms_list", multiple=True)
@click.pass_obj
def verify(
    ctx: Context,
    customers: tuple[str, ...],
    vendors: tuple[str, ...],
    items: tuple[str, ...],
    terms_list: tuple[str, ...],
) -> None:
    """Check whether referenced entities exist in QuickBooks."""
    result = _verify(
        ctx,
        customers=list(customers),
        vendors=list(vendors),
        items=list(items),
        terms=list(terms_list),
    )
    ok = True
    for bucket, missing in result.missing.items():
        if missing:
            ok = False
            for name in sorted(missing):
                click.echo(f"MISSING {bucket}: {name}", err=True)
    for bucket, found in result.found.items():
        for name in sorted(found):
            click.echo(f"found {bucket}: {name}")
    if not ok:
        raise click.exceptions.Exit(1)
