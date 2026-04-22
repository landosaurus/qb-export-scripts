from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal, Sequence

from qb_cli.context import Context
from qb_cli.io.csv_serializer import from_csv
from qb_cli.io.format import Format, detect_format
from qb_cli.io.json_serializer import from_json
from qb_cli.models.base import BaseEntity
from qb_cli.ops.dedupe import find_duplicates
from qb_cli.ops.registry import EntityHandler, get_handler
from qb_cli.ops.verify_op import verify_entities
from qb_cli.transport.errors import QBEntityNotFound, QBStatusError
from qb_cli.transport.status import check_response_status


OnDuplicate = Literal["error", "skip", "update"]


@dataclass
class ImportResult:
    entity: str
    attempted: int = 0
    written: int = 0
    skipped: int = 0
    failed: int = 0
    missing_refs: list[str] = field(default_factory=list)
    duplicate_refs: list[str] = field(default_factory=list)
    errors: list[tuple[str, str]] = field(default_factory=list)


def _load_entities(
    handler: EntityHandler, path: Path, fmt: Format
) -> list[BaseEntity]:
    if fmt is Format.CSV:
        return list(from_csv(handler.model, path))
    return list(from_json(handler.model, path))


def _ref_name(entity: BaseEntity, field_name: str) -> str | None:
    ref = getattr(entity, field_name, None)
    if ref is None:
        return None
    name = getattr(ref, "full_name", None)
    if isinstance(name, str) and name.strip():
        return name
    return None


def _collect_references(
    entity_key: str, entities: Sequence[BaseEntity]
) -> tuple[set[str], set[str], set[str], set[str]]:
    customers: set[str] = set()
    vendors: set[str] = set()
    items: set[str] = set()
    terms: set[str] = set()

    for e in entities:
        if entity_key == "purchase_order":
            vendor = _ref_name(e, "vendor_ref")
            if vendor:
                vendors.add(vendor)
        else:
            customer = _ref_name(e, "customer_ref")
            if customer:
                customers.add(customer)

        terms_name = _ref_name(e, "terms_ref")
        if terms_name:
            terms.add(terms_name)

        for line in getattr(e, "line_items", []) or []:
            item_ref = getattr(line, "item_ref", None)
            if item_ref is None:
                continue
            item_name = getattr(item_ref, "full_name", None)
            if isinstance(item_name, str) and item_name.strip():
                items.add(item_name)

    return customers, vendors, items, terms


def _build_missing_prefixed(missing: dict[str, set[str]]) -> list[str]:
    out: list[str] = []
    for bucket in sorted(missing.keys()):
        for name in sorted(missing[bucket]):
            out.append(f"{bucket}:{name}")
    return out


def _query_edit_sequence(
    ctx: Context, handler: EntityHandler, ref_number: str
) -> tuple[str | None, str | None]:
    request_xml = handler.qbxml.build_query(
        ref_numbers=[ref_number],
        include_line_items=False,
    )
    try:
        with ctx.connection_factory() as conn:
            response_xml = conn.send(request_xml)
        check_response_status(response_xml)
    except QBEntityNotFound:
        return None, None
    records = handler.qbxml.parse_query_response(response_xml)
    if not records:
        return None, None
    record = records[0]
    txn_id = getattr(record, "txn_id", None)
    edit_seq = getattr(record, "edit_sequence", None)
    return (
        txn_id if isinstance(txn_id, str) else None,
        edit_seq if isinstance(edit_seq, str) else None,
    )


def import_(
    ctx: Context,
    entity_key: str,
    *,
    input_path: str | Path,
    fmt: Format | None = None,
    dry_run: bool = False,
    on_duplicate: OnDuplicate = "error",
) -> ImportResult:
    handler = get_handler(entity_key)
    path = Path(input_path)
    resolved_fmt = fmt if fmt is not None else detect_format(path)

    entities = _load_entities(handler, path, resolved_fmt)
    result = ImportResult(entity=entity_key)

    customers, vendors, items, terms = _collect_references(entity_key, entities)

    verify = verify_entities(
        ctx,
        customers=sorted(customers),
        vendors=sorted(vendors),
        items=sorted(items),
        terms=sorted(terms),
    )
    if not verify.all_found:
        result.missing_refs = _build_missing_prefixed(verify.missing)
        if dry_run:
            result.attempted = len(entities)
        return result

    ref_numbers: list[str] = []
    for e in entities:
        ref = getattr(e, handler.id_field, None)
        if isinstance(ref, str):
            ref_numbers.append(ref)

    duplicates = find_duplicates(ctx, entity_key, ref_numbers)

    if duplicates and on_duplicate == "error":
        result.duplicate_refs = sorted(duplicates)
        return result

    planned_adds: list[BaseEntity] = []
    planned_mods: list[BaseEntity] = []
    skipped_count = 0

    for e in entities:
        ref = getattr(e, handler.id_field, None)
        ref_str = ref if isinstance(ref, str) else None
        if ref_str is not None and ref_str in duplicates:
            if on_duplicate == "skip":
                skipped_count += 1
                continue
            if on_duplicate == "update":
                if not handler.supports_mod:
                    result.errors.append(
                        (ref_str, f"{entity_key} does not support mod")
                    )
                    continue
                txn_id, edit_seq = _query_edit_sequence(ctx, handler, ref_str)
                if txn_id is None or edit_seq is None:
                    result.errors.append(
                        (ref_str, "could not resolve edit_sequence for update")
                    )
                    continue
                e.txn_id = txn_id
                e.edit_sequence = edit_seq
                planned_mods.append(e)
                continue
        planned_adds.append(e)

    result.skipped = skipped_count
    result.duplicate_refs = sorted(duplicates)
    attempted = len(planned_adds) + len(planned_mods)
    result.attempted = attempted

    if dry_run:
        return result

    for entity in planned_adds:
        ref_str = getattr(entity, handler.id_field, None)
        ref_display = ref_str if isinstance(ref_str, str) else ""
        try:
            request_xml = handler.qbxml.build_add(entity)
            with ctx.connection_factory() as conn:
                response_xml = conn.send(request_xml)
            check_response_status(response_xml)
            handler.qbxml.parse_add_response(response_xml)
            result.written += 1
        except QBStatusError as e:
            result.failed += 1
            result.errors.append((ref_display, str(e)))

    for entity in planned_mods:
        ref_str = getattr(entity, handler.id_field, None)
        ref_display = ref_str if isinstance(ref_str, str) else ""
        try:
            request_xml = handler.qbxml.build_mod(entity)
            with ctx.connection_factory() as conn:
                response_xml = conn.send(request_xml)
            check_response_status(response_xml)
            handler.qbxml.parse_mod_response(response_xml)
            result.written += 1
        except QBStatusError as e:
            result.failed += 1
            result.errors.append((ref_display, str(e)))

    return result
