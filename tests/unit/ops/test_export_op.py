from __future__ import annotations

import json
import logging
from pathlib import Path
from types import SimpleNamespace

import pytest

from qb_cli.io.format import Format
from qb_cli.ops.export_op import ExportResult, export
from tests.unit.ops._fakes import FakeConnection, make_factory


_FIXTURE_DIR = Path(__file__).resolve().parents[2] / "fixtures" / "qbxml_responses"
_NULL_LOGGER = logging.getLogger("qb_cli.test")
_NULL_LOGGER.addHandler(logging.NullHandler())


def _ctx_with(response_xml: str) -> SimpleNamespace:
    return SimpleNamespace(
        connection_factory=make_factory(response_xml),
        logger=_NULL_LOGGER,
    )


def test_export_invoice_to_csv(tmp_path: Path):
    fixture = (_FIXTURE_DIR / "invoice_query_sample.xml").read_text()
    ctx = _ctx_with(fixture)
    out = tmp_path / "inv.csv"

    result = export(ctx, "invoice", output_path=out)

    assert isinstance(result, ExportResult)
    assert result.count == 2
    assert result.output_path == out
    assert result.format is Format.CSV
    assert out.exists()
    assert "row_type" in out.read_text().splitlines()[0]


def test_export_sales_order_to_json(tmp_path: Path):
    fixture = (_FIXTURE_DIR / "sales_order_query_sample.xml").read_text()
    ctx = _ctx_with(fixture)
    out = tmp_path / "so.json"

    result = export(ctx, "sales_order", output_path=out)

    assert result.count == 2
    assert result.format is Format.JSON
    data = json.loads(out.read_text())
    assert isinstance(data, list)
    assert len(data) == 2


def test_export_purchase_order_sends_query_with_filters(tmp_path: Path):
    fixture = (_FIXTURE_DIR / "purchase_order_query_sample.xml").read_text()
    factory = make_factory(fixture)
    ctx = SimpleNamespace(connection_factory=factory, logger=_NULL_LOGGER)
    out = tmp_path / "po.json"

    result = export(
        ctx,
        "purchase_order",
        ref_numbers=["7740", "7741"],
        output_path=out,
    )
    assert result.count == 2
    assert result.pages == 1  # ref-filtered queries never iterate


def test_export_unknown_entity_raises(tmp_path: Path):
    ctx = _ctx_with("")
    with pytest.raises(KeyError):
        export(ctx, "widget", output_path=tmp_path / "out.csv")


def _invoice_page(rows_xml: str, *, iterator_id: str | None = None, remaining: int = 0) -> str:
    attrs = f'statusCode="0" statusSeverity="Info" statusMessage="Status OK"'
    if iterator_id is not None:
        attrs += f' iteratorID="{iterator_id}" iteratorRemainingCount="{remaining}"'
    return f"""<?xml version="1.0" ?>
<QBXML><QBXMLMsgsRs><InvoiceQueryRs requestID="1" {attrs}>
{rows_xml}
</InvoiceQueryRs></QBXMLMsgsRs></QBXML>
"""


def _invoice_ret(ref_number: str) -> str:
    return f"""<InvoiceRet>
  <TxnID>T-{ref_number}</TxnID>
  <EditSequence>1</EditSequence>
  <CustomerRef><FullName>ACME</FullName></CustomerRef>
  <RefNumber>{ref_number}</RefNumber>
</InvoiceRet>"""


def test_export_iterates_across_pages(tmp_path: Path):
    page1 = _invoice_page(
        "\n".join(_invoice_ret(str(n)) for n in range(1, 4)),
        iterator_id="ITER-1",
        remaining=2,
    )
    page2 = _invoice_page(
        "\n".join(_invoice_ret(str(n)) for n in range(4, 6)),
        iterator_id="ITER-1",
        remaining=0,
    )
    # Use a stateful factory so both pages are served from the same FakeConnection instance.
    connection = FakeConnection(responses=[page1, page2])
    ctx = SimpleNamespace(connection_factory=lambda: connection, logger=_NULL_LOGGER)
    out = tmp_path / "iter.csv"

    result = export(ctx, "invoice", output_path=out, page_size=3)

    assert result.count == 5
    assert result.pages == 2


def test_export_iterator_second_request_carries_iterator_id(tmp_path: Path):
    page1 = _invoice_page(_invoice_ret("1"), iterator_id="ITER-X", remaining=1)
    page2 = _invoice_page(_invoice_ret("2"), iterator_id="ITER-X", remaining=0)

    connection = FakeConnection(responses=[page1, page2])

    def factory() -> FakeConnection:
        return connection

    ctx = SimpleNamespace(connection_factory=factory, logger=_NULL_LOGGER)
    out = tmp_path / "iter2.csv"
    result = export(ctx, "invoice", output_path=out, page_size=1)

    assert result.count == 2
    assert result.pages == 2
    assert len(connection.sent_requests) == 2
    assert 'iterator="Start"' in connection.sent_requests[0]
    assert 'iterator="Continue"' in connection.sent_requests[1]
    assert 'iteratorID="ITER-X"' in connection.sent_requests[1]


def test_export_single_page_when_no_iteratorid_in_response(tmp_path: Path):
    # Response has no iteratorID attribute → treat as single page even without ref filter.
    page = _invoice_page(_invoice_ret("1"))
    ctx = _ctx_with(page)
    out = tmp_path / "single.csv"
    result = export(ctx, "invoice", output_path=out)
    assert result.count == 1
    assert result.pages == 1
