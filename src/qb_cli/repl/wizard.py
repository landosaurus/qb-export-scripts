from __future__ import annotations

from datetime import date
from enum import Enum
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

# Navigation sentinels
_BACK_LABEL = "← Back"
_MAIN_LABEL = "← Main menu"


class Nav(Enum):
    """Sentinel values returned by prompt helpers to signal navigation."""

    BACK = "back"
    MAIN = "main"


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


# ---------- Prompt helpers with navigation support ----------


def _select(
    message: str,
    choices: list[str],
    *,
    include_back: bool = True,
) -> str | Nav | None:
    """Wrap questionary.select with nav choices and numbered hotkeys.

    Returns:
        - The user's choice string (one of ``choices``) on selection
        - Nav.BACK if they selected the Back option
        - Nav.MAIN if they selected the Main menu option
        - None if they cancelled (Ctrl-C / ESC)
    """
    display_choices: list[str] = []
    if include_back:
        display_choices.append(_BACK_LABEL)
    display_choices.append(_MAIN_LABEL)
    display_choices.extend(choices)
    result = questionary.select(
        message, choices=display_choices, use_shortcuts=True
    ).ask()
    if result is None:
        return None
    if result == _BACK_LABEL:
        return Nav.BACK
    if result == _MAIN_LABEL:
        return Nav.MAIN
    return result


def _text(message: str, *, default: str = "") -> str | Nav | None:
    """Wrap questionary.text with 'b'/'m' sentinel support.

    Returns Nav.BACK / Nav.MAIN / None as appropriate, else the raw string.
    """
    hinted = f"{message} (b=back, m=main)"
    if default:
        result = questionary.text(hinted, default=default).ask()
    else:
        result = questionary.text(hinted).ask()
    if result is None:
        return None
    stripped = result.strip().lower()
    if stripped in ("b", "back"):
        return Nav.BACK
    if stripped in ("m", "main", "main menu"):
        return Nav.MAIN
    return result


def _path(message: str) -> str | Nav | None:
    """Wrap questionary.path with 'b'/'m' sentinel support."""
    hinted = f"{message} (b=back, m=main)"
    result = questionary.path(hinted).ask()
    if result is None:
        return None
    stripped = result.strip().lower()
    if stripped in ("b", "back"):
        return Nav.BACK
    if stripped in ("m", "main", "main menu"):
        return Nav.MAIN
    return result


def _confirm(message: str, *, default: bool = True) -> bool | None:
    """Wrap questionary.confirm. No sentinels; Ctrl-C returns None -> main menu."""
    return questionary.confirm(message, default=default).ask()


def _checkbox(message: str, choices: list[str]) -> list[str] | Nav | None:
    """Wrap questionary.checkbox with a single 'Main menu' escape choice.

    If the user ticks the main-menu choice (and confirms), we return Nav.MAIN.
    ``Back`` is not offered on checkbox prompts (the spec).
    """
    display = [_MAIN_LABEL] + choices
    result = questionary.checkbox(message, choices=display).ask()
    if result is None:
        return None
    if _MAIN_LABEL in result:
        return Nav.MAIN
    return result


# ---------- Entity / format / mode prompts ----------


def _ask_entity(
    message: str = "Entity type:", *, include_back: bool = True
) -> str | Nav | None:
    return _select(message, _ordered_entity_choices(), include_back=include_back)


def _ask_export_mode() -> str | Nav | None:
    return _select(
        "How do you want to select records?",
        [_MODE_REFS, _MODE_YEAR, _MODE_DATE_RANGE],
    )


def _ask_format(message: str = "Output format:") -> Format | Nav | None:
    label = _select(message, ["CSV", "JSON"])
    if label is None or isinstance(label, Nav):
        return label
    return Format.CSV if label == "CSV" else Format.JSON


# ---------- Export flow ----------


def _run_export(ctx: Context) -> None:
    """Export sub-flow, modeled as a step machine.

    Each step fn reads/writes from ``state`` and returns either:
      - None — user cancelled (Ctrl-C) -> return to main menu
      - Nav.BACK — go to previous step (no-op at step 0)
      - Nav.MAIN — return to main menu
      - any other value — advance to next step
    The final step runs the actual export and returns a non-Nav value.
    """

    state: dict[str, object] = {}

    def step_entity() -> str | Nav | None:
        # First step of sub-flow: Back == Main menu, so hide Back.
        result = _ask_entity(include_back=False)
        if isinstance(result, str):
            state["entity_key"] = result
        return result

    def step_mode() -> str | Nav | None:
        result = _ask_export_mode()
        if isinstance(result, str):
            state["mode"] = result
        return result

    def step_mode_details() -> str | Nav | None:
        mode = state["mode"]
        assert isinstance(mode, str)
        if mode == _MODE_REFS:
            raw = _text("Ref numbers (comma-separated):")
            if raw is None or isinstance(raw, Nav):
                return raw
            state["ref_numbers"] = _split_refs(raw)
            state["date_from"] = None
            state["date_to"] = None
            return "ok"
        if mode == _MODE_YEAR:
            raw_year = _text("Year (YYYY):")
            if raw_year is None or isinstance(raw_year, Nav):
                return raw_year
            try:
                year = int(raw_year.strip())
            except ValueError:
                click.echo(f"error: invalid year: {raw_year!r}", err=True)
                return Nav.MAIN
            date_from, date_to = year_range(year)
            state["ref_numbers"] = None
            state["date_from"] = date_from
            state["date_to"] = date_to
            return "ok"
        # _MODE_DATE_RANGE
        raw_from = _text("From date (YYYY-MM-DD):")
        if raw_from is None or isinstance(raw_from, Nav):
            return raw_from
        raw_to = _text("To date (YYYY-MM-DD):")
        if raw_to is None or isinstance(raw_to, Nav):
            return raw_to
        try:
            state["ref_numbers"] = None
            state["date_from"] = parse_date(raw_from)
            state["date_to"] = parse_date(raw_to)
        except ValueError as e:
            click.echo(f"error: {e}", err=True)
            return Nav.MAIN
        return "ok"

    def step_format() -> Format | Nav | None:
        fmt = _ask_format()
        if isinstance(fmt, Format):
            state["fmt"] = fmt
        return fmt

    def step_output_and_run() -> str | Nav | None:
        entity_key = state["entity_key"]
        fmt = state["fmt"]
        assert isinstance(entity_key, str)
        assert isinstance(fmt, Format)
        suggested = _suggest_output_path(entity_key, fmt)
        path_raw = _text("Output file:", default=suggested)
        if path_raw is None or isinstance(path_raw, Nav):
            return path_raw
        output_path = Path(path_raw)
        ref_numbers = state.get("ref_numbers")
        date_from = state.get("date_from")
        date_to = state.get("date_to")
        assert ref_numbers is None or isinstance(ref_numbers, list)
        assert date_from is None or isinstance(date_from, date)
        assert date_to is None or isinstance(date_to, date)
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
            return "done"
        click.echo(
            f"exported {result.count} records to {result.output_path} (pages={result.pages})"
        )
        return "done"

    steps: list[Callable[[], object]] = [
        step_entity,
        step_mode,
        step_mode_details,
        step_format,
        step_output_and_run,
    ]
    _drive_flow(steps)


def _drive_flow(steps: list[Callable[[], object]]) -> None:
    """Drive a list of step functions with Back/Main-menu semantics.

    - None or Nav.MAIN -> return (back to main menu)
    - Nav.BACK         -> decrement index (clamped at 0)
    - anything else    -> advance to next step
    - after the last step, return
    """
    idx = 0
    while idx < len(steps):
        result = steps[idx]()
        if result is None or result is Nav.MAIN:
            return
        if result is Nav.BACK:
            if idx == 0:
                # Safe fallback: Back at the first step == Main menu.
                return
            idx -= 1
            continue
        idx += 1


# ---------- Import flow ----------


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


def _to_on_duplicate(value: str) -> OnDuplicate:
    """Narrow a user-supplied string to the OnDuplicate literal type."""
    if value == "skip":
        return "skip"
    if value == "update":
        return "update"
    return "error"


def _run_import(ctx: Context) -> None:
    state: dict[str, object] = {}

    def step_entity() -> str | Nav | None:
        result = _ask_entity(include_back=False)
        if isinstance(result, str):
            state["entity_key"] = result
        return result

    def step_path() -> str | Nav | None:
        result = _path("Input file:")
        if result is None or isinstance(result, Nav):
            return result
        path = Path(result)
        state["input_path"] = path
        detected = _detect_import_format(path)
        state["detected_fmt"] = detected
        return "ok"

    def step_format() -> Format | Nav | None | str:
        detected = state.get("detected_fmt")
        if isinstance(detected, Format):
            state["fmt"] = detected
            return "ok"
        fmt = _ask_format("Input format (could not auto-detect):")
        if isinstance(fmt, Format):
            state["fmt"] = fmt
        return fmt

    def step_dry_run() -> bool | Nav | None:
        answer = _confirm("Dry run? (recommended for first import)", default=True)
        if answer is None:
            return None
        state["dry_run"] = answer
        return answer

    def step_on_duplicate() -> str | Nav | None:
        result = _select("On duplicate:", ["error", "skip", "update"])
        if isinstance(result, str):
            state["on_duplicate"] = _to_on_duplicate(result)
        return result

    def step_confirm_and_run() -> str | Nav | None:
        entity_key = state["entity_key"]
        input_path = state["input_path"]
        fmt = state["fmt"]
        dry_run = state["dry_run"]
        on_duplicate = state["on_duplicate"]
        assert isinstance(entity_key, str)
        assert isinstance(input_path, Path)
        assert isinstance(fmt, Format)
        assert isinstance(dry_run, bool)
        assert isinstance(on_duplicate, str)
        summary = (
            f"About to import {entity_key} from {input_path} "
            f"(format={fmt.value}, dry_run={dry_run}, on_duplicate={on_duplicate})"
        )
        click.echo(summary)
        proceed = _confirm("Proceed?", default=True)
        if proceed is None:
            return None
        if not proceed:
            return "done"
        on_dup: OnDuplicate = _to_on_duplicate(on_duplicate)
        try:
            result = import_(
                ctx,
                entity_key,
                input_path=input_path,
                fmt=fmt,
                dry_run=dry_run,
                on_duplicate=on_dup,
            )
        except Exception as e:  # noqa: BLE001 — user-facing top-level guard
            click.echo(f"error: import failed: {e}", err=True)
            return "done"
        _echo_import_result(result)
        return "done"

    steps: list[Callable[[], object]] = [
        step_entity,
        step_path,
        step_format,
        step_dry_run,
        step_on_duplicate,
        step_confirm_and_run,
    ]
    _drive_flow(steps)


# ---------- Query flow ----------


def _format_query_row(entity: BaseEntity) -> str:
    ref = getattr(entity, "ref_number", None)
    txn_date = getattr(entity, "txn_date", None)
    total = getattr(entity, "total_amount", None)
    ref_s = str(ref) if ref is not None else "-"
    date_s = str(txn_date) if txn_date is not None else "-"
    total_s = str(total) if total is not None else "-"
    return f"{ref_s}\t{date_s}\t{total_s}"


def _run_query(ctx: Context) -> None:
    state: dict[str, object] = {}

    def step_entity() -> str | Nav | None:
        result = _ask_entity(include_back=False)
        if isinstance(result, str):
            state["entity_key"] = result
        return result

    def step_refs_and_run() -> str | Nav | None:
        raw = _text("Ref numbers:")
        if raw is None or isinstance(raw, Nav):
            return raw
        refs = _split_refs(raw)
        entity_key = state["entity_key"]
        assert isinstance(entity_key, str)
        try:
            results = query(ctx, entity_key, ref_numbers=refs if refs else None)
        except Exception as e:  # noqa: BLE001 — user-facing top-level guard
            click.echo(f"error: query failed: {e}", err=True)
            return "done"
        for entity in results:
            click.echo(_format_query_row(entity))
        return "done"

    steps: list[Callable[[], object]] = [step_entity, step_refs_and_run]
    _drive_flow(steps)


# ---------- Verify flow ----------


_BUCKET_TO_SINGULAR: dict[str, str] = {
    "customers": "customer",
    "vendors": "vendor",
    "items": "item",
    "terms": "terms",
}


def _run_verify(ctx: Context) -> None:
    state: dict[str, object] = {}

    def step_buckets() -> list[str] | Nav | None:
        result = _checkbox(
            "Which entity types?",
            ["customers", "vendors", "items", "terms"],
        )
        if result is None or isinstance(result, Nav):
            return result
        if not result:
            click.echo("no buckets selected")
            return Nav.MAIN
        state["selected"] = result
        return result

    def step_refs_and_run() -> str | Nav | None:
        selected = state["selected"]
        assert isinstance(selected, list)
        bucket_kwargs: dict[str, list[str]] = {}
        for bucket in selected:
            raw = _text(f"{bucket} (comma-separated):")
            if raw is None or isinstance(raw, Nav):
                return raw
            bucket_kwargs[bucket] = _split_refs(raw)
        try:
            result = _call_verify(ctx, bucket_kwargs)
        except Exception as e:  # noqa: BLE001 — user-facing top-level guard
            click.echo(f"error: verify failed: {e}", err=True)
            return "done"
        for bucket_plural in ("customers", "vendors", "items", "terms"):
            singular = _BUCKET_TO_SINGULAR[bucket_plural]
            for name in sorted(result.found.get(singular, set())):
                click.echo(f"found {bucket_plural}: {name}")
            for name in sorted(result.missing.get(singular, set())):
                click.echo(f"MISSING {bucket_plural}: {name}", err=True)
        return "done"

    steps: list[Callable[[], object]] = [step_buckets, step_refs_and_run]
    _drive_flow(steps)


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


# ---------- Main menu / entry point ----------


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
        use_shortcuts=True,
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
