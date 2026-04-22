from decimal import Decimal
from pathlib import Path
from qb_cli.qbxml.invoice import parse_query_response


FIXTURE = Path(__file__).parent.parent.parent / "fixtures" / "qbxml_responses" / "invoice_query_sample.xml"


def test_parse_query_returns_two_invoices():
    xml = FIXTURE.read_text()
    records = parse_query_response(xml)
    assert len(records) == 2
    assert records[0].ref_number == "14396"
    assert records[1].ref_number == "14397"


def test_parse_first_invoice_header_fields():
    xml = FIXTURE.read_text()
    records = parse_query_response(xml)
    r0 = records[0]
    assert r0.txn_id == "ABC-1"
    assert r0.customer_ref.full_name == "ACME Corp"
    assert r0.po_number == "7740-SH"
    assert r0.ship_address.city == "Seattle"


def test_parse_second_invoice_line_items():
    xml = FIXTURE.read_text()
    records = parse_query_response(xml)
    r1 = records[1]
    assert len(r1.line_items) == 2
    assert r1.line_items[0].item_ref.full_name == "40-RAG12"
    assert r1.line_items[0].amount == Decimal("250.00")
    assert r1.line_items[1].item_ref.full_name == "Freight"


def test_parse_empty_line_items_ok():
    xml = FIXTURE.read_text()
    records = parse_query_response(xml)
    assert records[0].line_items == []
