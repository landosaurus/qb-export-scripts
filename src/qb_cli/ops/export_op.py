from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Sequence

from qb_cli.context import Context
from qb_cli.io.csv_serializer import to_csv
from qb_cli.io.format import Format, detect_format
from qb_cli.io.json_serializer import to_json
from qb_cli.ops.registry import get_handler
from qb_cli.transport.status import check_response_status


@dataclass
class ExportResult:
    entity: str
    count: int
    output_path: Path
    format: Format


def export(
    ctx: Context,
    entity_key: str,
    *,
    ref_numbers: Sequence[str] | None = None,
    date_from: date | None = None,
    date_to: date | None = None,
    output_path: str | Path,
    fmt: Format | None = None,
) -> ExportResult:
    handler = get_handler(entity_key)
    path = Path(output_path)
    resolved_fmt = fmt if fmt is not None else detect_format(path)

    request_xml = handler.qbxml.build_query(
        ref_numbers=list(ref_numbers) if ref_numbers else None,
        date_from=date_from,
        date_to=date_to,
        include_line_items=True,
    )
    with ctx.connection_factory() as conn:
        response_xml = conn.send(request_xml)
    check_response_status(response_xml)
    records = handler.qbxml.parse_query_response(response_xml)

    if resolved_fmt is Format.CSV:
        to_csv(records, path)
    else:
        to_json(records, path)

    return ExportResult(
        entity=entity_key, count=len(records), output_path=path, format=resolved_fmt
    )
