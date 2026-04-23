from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Callable

import click
import questionary

from qb_cli.context import Context
from qb_cli.io.format import Format, detect_format
from qb_cli.models.base import BaseEntity
from qb_cli.ops.export_op import export
from qb_cli.ops.import_op import ImportResult, OnDuplicate, import_
from qb_cli.ops.query_op import query
from qb_cli.ops.registry import HANDLERS
from qb_cli.ops.verify_op import VerifyResult, verify_entities
from qb_cli.utils.date_parse import parse_date, year_range


# Main menu labels
_MAIN_EXPORT = "Export records"
_MAIN_IMPORT = "Import records"
_MAIN_QUERY = "Query records"
_MAIN_VERIFY = "Verify entities exist"
_MAIN_EXIT = "Exit"

# Export mode labels
_MODE_REFS = "Specific ref numbers (comma-separated)"
_MODE_YEAR = "Year (all records from Jan 1 through Dec 31)"
_MODE_DATE_RANGE = "Date range"


def _ordered_entity_choices() -> list[str]:
    """Return entity keys with invoice, sales_order, purchase_order first (fixed order),
    then any others alphabetically."""
    preferred: list[str] = ["invoice", "sales_order", "purchase_order"]
    known = set(HANDLERS.keys())
    out: list[str] = [k for k in preferred if k in known]
    extras = sorted(k for k in known if k not in preferred)
    out.extend(extras)
    return out


def _split_refs(raw: str) -> list[str]:
    """Split a comma-separated string into a clean list of refs; drop empties."""
    return [x.strip() for x in raw.split(",") if x.strip()]


def _suggest_output_path(entity_key: str, fmt: Format) -> str:
    today = date.today().isoformat()
    ext = "csv" if fmt is Format.CSV else "json"
    return f"{entity_key}_{today}.{ext}"


def _ask_entity(message: str = "Entity type:") -> str | None:
    return questionary.select(message, choices=_ordered_entity_choices()).ask()


def _ask_export_mode() -> str | None:
    return questionary.select(
        "How do you want to select records?",
        choices=[_MODE_REFS, _MODE_YEAR, _MODE_DATE_RANGE],
    ).ask()


def _ask_format(message: str = "Output format:") -> Format | None:
    label = questionary.select(message, choices=["CSV", "JSON"]).ask()
    if label is None:
        return None
    return Format.CSV if label == "CSV" else Format.JSON


def _run_export(ctx: Context) -> None:
    entity_key = _ask_entity()
    if entity_key is None:
        return

    mode = _ask_export_mode()
    if mode is None:
        return

    ref_numbers: list[str] | None = None
    date_from: date | None = None
    date_to: date | None = None

    if mode == _MODE_REFS:
        raw = questionary.text("Ref numbers (comma-separated):").ask()
        if raw is None:
            return
        ref_numbers = _split_refs(raw)
    elif mode == _MODE_YEAR:
        raw_year = questionary.text("Year (YYYY):").ask()
        if raw_year is None:
            return
        try:
            year = int(raw_year.strip())
        except ValueError:
            click.echo(f"error: invalid year: {raw_year!r}", err=True)
            return
        date_from, date_to = year_range(year)
    else:  # _MODE_DATE_RANGE
        raw_from = questionary.text("From date (YYYY-MM-DD):").ask()
        if raw_from is None:
            return
        raw_to = questionary.text("To date (YYYY-MM-DD):").ask()
        if raw_to is None:
            return
        try:
            date_from = parse_date(raw_from)
            date_to = parse_date(raw_to)
        except ValueError as e:
            click.echo(f"error: {e}", err=True)
            return

    fmt = _ask_format()
    if fmt is None:
        return

    suggested = _suggest_output_path(entity_key, fmt)
    path_raw = questionary.text("Output file:", default=suggested).ask()
    if path_raw is None:
        return
    output_path = Path(path_raw)

    try:
        result = export(
            ctx,
            entity_key,
            ref_numbers=ref_numbers,
            date_from=date_from,
            date_to=date_to,
            output_path=output_path,
            fmt=fmt,
        )
    except Exception as e:  # noqa: BLE001 — user-facing top-level guard
        click.echo(f"error: export failed: {e}", err=True)
        return

    click.echo(
        f"exported {result.count} records to {result.output_path} (pages={result.pages})"
    )


def _detect_import_format(path: Path) -> Format | None:
    """Try to auto-detect format from the path suffix. Return None if ambiguous."""
    try:
        return detect_format(path)
    except ValueError:
        return None


def _echo_import_result(result: ImportResult) -> None:
    click.echo(
        f"import: attempted={result.attempted} written={result.written} "
        f"skipped={result.skipped} failed={result.failed}"
    )
    if result.missing_refs:
        click.echo("missing refs:", err=True)
        for ref in result.missing_refs:
            click.echo(f"  {ref}", err=True)
    if result.duplicate_refs:
        click.echo("duplicate refs:", err=True)
        for ref in result.duplicate_refs:
            click.echo(f"  {ref}", err=True)
    if result.errors:
        click.echo("errors:", err=True)
        for ref, msg in result.errors:
            click.echo(f"  {ref}: {msg}", err=True)


def _run_import(ctx: Context) -> None:
    entity_key = _ask_entity()
    if entity_key is None:
        return

    path_raw = questionary.path("Input file:").ask()
    if path_raw is None:
        return
    input_path = Path(path_raw)

    fmt = _detect_import_format(input_path)
    if fmt is None:
        chosen = _ask_format("Input format (could not auto-detect):")
        if chosen is None:
            return
        fmt = chosen

    dry_run = questionary.confirm(
        "Dry run? (recommended for first import)", default=True
    ).ask()
    if dry_run is None:
        return

    on_dup_raw = questionary.select(
        "On duplicate:",
        choices=["error", "skip", "update"],
        default="error",
    ).ask()
    if on_dup_raw is None:
        return
    on_duplicate: OnDuplicate = _to_on_duplicate(on_dup_raw)

    summary = (
        f"About to import {entity_key} from {input_path} "
        f"(format={fmt.value}, dry_run={dry_run}, on_duplicate={on_duplicate})"
    )
    click.echo(summary)
    proceed = questionary.confirm("Proceed?", default=True).ask()
    if not proceed:
        return

    try:
        result = import_(
            ctx,
            entity_key,
            input_path=input_path,
            fmt=fmt,
            dry_run=dry_run,
            on_duplicate=on_duplicate,
        )
    except Exception as e:  # noqa: BLE001 — user-facing top-level guard
        click.echo(f"error: import failed: {e}", err=True)
        return

    _echo_import_result(result)


def _to_on_duplicate(value: str) -> OnDuplicate:
    """Narrow a user-supplied string to the OnDuplicate literal type."""
    if value == "skip":
        return "skip"
    if value == "update":
        return "update"
    return "error"


def _format_query_row(entity: BaseEntity) -> str:
    ref = getattr(entity, "ref_number", None)
    txn_date = getattr(entity, "txn_date", None)
    total = getattr(entity, "total_amount", None)
    ref_s = str(ref) if ref is not None else "-"
    date_s = str(txn_date) if txn_date is not None else "-"
    total_s = str(total) if total is not None else "-"
    return f"{ref_s}\t{date_s}\t{total_s}"


def _run_query(ctx: Context) -> None:
    entity_key = _ask_entity()
    if entity_key is None:
        return

    raw = questionary.text("Ref numbers:").ask()
    if raw is None:
        return
    refs = _split_refs(raw)

    try:
        results = query(ctx, entity_key, ref_numbers=refs if refs else None)
    except Exception as e:  # noqa: BLE001 — user-facing top-level guard
        click.echo(f"error: query failed: {e}", err=True)
        return

    for entity in results:
        click.echo(_format_query_row(entity))


_BUCKET_TO_SINGULAR: dict[str, str] = {
    "customers": "customer",
    "vendors": "vendor",
    "items": "item",
    "terms": "terms",
}


def _run_verify(ctx: Context) -> None:
    selected = questionary.checkbox(
        "Which entity types?",
        choices=["customers", "vendors", "items", "terms"],
    ).ask()
    if selected is None:
        return
    if not selected:
        click.echo("no buckets selected")
        return

    bucket_kwargs: dict[str, list[str]] = {}
    for bucket in selected:
        raw = questionary.text(f"{bucket} (comma-separated):").ask()
        if raw is None:
            return
        bucket_kwargs[bucket] = _split_refs(raw)

    try:
        result = _call_verify(ctx, bucket_kwargs)
    except Exception as e:  # noqa: BLE001 — user-facing top-level guard
        click.echo(f"error: verify failed: {e}", err=True)
        return

    for bucket_plural in ("customers", "vendors", "items", "terms"):
        singular = _BUCKET_TO_SINGULAR[bucket_plural]
        for name in sorted(result.found.get(singular, set())):
            click.echo(f"found {bucket_plural}: {name}")
        for name in sorted(result.missing.get(singular, set())):
            click.echo(f"MISSING {bucket_plural}: {name}", err=True)


def _call_verify(
    ctx: Context, bucket_kwargs: dict[str, list[str]]
) -> VerifyResult:
    """Wrapper that translates the (plural-keyed) dict into the verify_entities kwargs.

    Each bucket that wasn't selected gets an empty tuple, per the spec.
    """
    return verify_entities(
        ctx,
        customers=bucket_kwargs.get("customers", ()) or (),
        vendors=bucket_kwargs.get("vendors", ()) or (),
        items=bucket_kwargs.get("items", ()) or (),
        terms=bucket_kwargs.get("terms", ()) or (),
    )


def _main_menu_choice() -> str | None:
    return questionary.select(
        "What would you like to do?",
        choices=[
            _MAIN_EXPORT,
            _MAIN_IMPORT,
            _MAIN_QUERY,
            _MAIN_VERIFY,
            _MAIN_EXIT,
        ],
    ).ask()


_ACTIONS: dict[str, Callable[[Context], None]] = {
    _MAIN_EXPORT: _run_export,
    _MAIN_IMPORT: _run_import,
    _MAIN_QUERY: _run_query,
    _MAIN_VERIFY: _run_verify,
}


def run_wizard(ctx: Context) -> int:
    """Launch the interactive wizard. Returns the process exit code.

    Ctrl-C / ESC at any prompt returns 0 (user cancelled).
    """
    while True:
        choice = _main_menu_choice()
        if choice is None or choice == _MAIN_EXIT:
            return 0
        action = _ACTIONS.get(choice)
        if action is None:
            # Unknown label — defensively fall back to exit to avoid infinite loop.
            return 0
        action(ctx)
