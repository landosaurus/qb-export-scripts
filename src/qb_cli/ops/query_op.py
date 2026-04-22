from __future__ import annotations

from typing import Sequence

from qb_cli.context import Context
from qb_cli.models.base import BaseEntity
from qb_cli.ops.registry import get_handler
from qb_cli.transport.status import check_response_status


def query(
    ctx: Context,
    entity_key: str,
    *,
    ref_numbers: Sequence[str] | None = None,
) -> list[BaseEntity]:
    handler = get_handler(entity_key)
    request_xml = handler.qbxml.build_query(
        ref_numbers=list(ref_numbers) if ref_numbers else None,
        include_line_items=True,
    )
    with ctx.connection_factory() as conn:
        response_xml = conn.send(request_xml)
    check_response_status(response_xml)
    return handler.qbxml.parse_query_response(response_xml)
