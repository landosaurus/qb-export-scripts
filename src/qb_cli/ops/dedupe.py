from __future__ import annotations

from typing import Sequence

from qb_cli.context import Context
from qb_cli.ops.registry import get_handler
from qb_cli.transport.errors import QBEntityNotFound
from qb_cli.transport.status import check_response_status


def find_duplicates(ctx: Context, entity_key: str, ref_numbers: Sequence[str]) -> set[str]:
    """Return the set of ref_numbers that already exist in QB for entity_key."""
    if not ref_numbers:
        return set()
    handler = get_handler(entity_key)
    request_xml = handler.qbxml.build_query(
        ref_numbers=list(ref_numbers),
        include_line_items=False,
    )
    try:
        with ctx.connection_factory() as conn:
            response_xml = conn.send(request_xml)
        check_response_status(response_xml)
    except QBEntityNotFound:
        return set()

    records = handler.qbxml.parse_query_response(response_xml)
    requested = set(ref_numbers)
    found: set[str] = set()
    for rec in records:
        ref_value = getattr(rec, handler.id_field, None)
        if isinstance(ref_value, str):
            found.add(ref_value)
    return found & requested
