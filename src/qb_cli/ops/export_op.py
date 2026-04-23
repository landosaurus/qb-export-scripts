from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Sequence

from qb_cli.context import Context
from qb_cli.io.csv_serializer import to_csv
from qb_cli.io.format import Format, detect_format
from qb_cli.io.json_serializer import to_json
from qb_cli.models.base import BaseEntity
from qb_cli.ops.registry import get_handler
from qb_cli.qbxml.common import read_iterator_state
from qb_cli.transport.status import check_response_status

DEFAULT_PAGE_SIZE = 500


@dataclass
class ExportResult:
    entity: str
    count: int
    output_path: Path
    format: Format
    pages: int = 1


def export(
    ctx: Context,
    entity_key: str,
    *,
    ref_numbers: Sequence[str] | None = None,
    date_from: date | None = None,
    date_to: date | None = None,
    output_path: str | Path,
    fmt: Format | None = None,
    page_size: int = DEFAULT_PAGE_SIZE,
) -> ExportResult:
    handler = get_handler(entity_key)
    path = Path(output_path)
    resolved_fmt = fmt if fmt is not None else detect_format(path)

    refs = list(ref_numbers) if ref_numbers else None
    # QBXML forbids MaxReturned / iterator when filtering by RefNumber or TxnID.
    # In that case we do exactly one request.
    paginate = refs is None

    records: list[BaseEntity] = []
    pages = 0
    iterator_id: str | None = None

    with ctx.connection_factory() as conn:
        while True:
            pages += 1
            request_xml = handler.qbxml.build_query(
                ref_numbers=refs,
                date_from=date_from,
                date_to=date_to,
                include_line_items=True,
                max_returned=page_size if paginate else None,
                iterator_id=iterator_id,
            )
            response_xml = conn.send(request_xml)
            check_response_status(response_xml)
            records.extend(handler.qbxml.parse_query_response(response_xml))
            ctx.logger.info(
                "exported page %d (%d records so far) for %s",
                pages, len(records), entity_key,
            )

            if not paginate:
                break
            iterator_id, remaining = read_iterator_state(response_xml)
            if not iterator_id or remaining <= 0:
                break

    if resolved_fmt is Format.CSV:
        to_csv(records, path)
    else:
        to_json(records, path)

    return ExportResult(
        entity=entity_key,
        count=len(records),
        output_path=path,
        format=resolved_fmt,
        pages=pages,
    )
