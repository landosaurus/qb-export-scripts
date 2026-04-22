from decimal import Decimal
from pathlib import Path
from qb_cli.qbxml.purchase_order import parse_query_response


FIXTURE = Path(__file__).parent.parent.parent / "fixtures" / "qbxml_responses" / "purchase_order_query_sample.xml"


def test_parse_query_returns_two_records():
    records = parse_query_response(FIXTURE.read_text())
    assert len(records) == 2
    assert records[0].ref_number == "7740"
    assert records[1].ref_number == "7741"


def test_parse_first_header_fields():
    records = parse_query_response(FIXTURE.read_text())
    r0 = records[0]
    assert r0.txn_id == "PO-1"
    assert r0.vendor_ref.full_name == "Widgets Inc."
    assert r0.vendor_address.city == "Portland"


def test_parse_second_line_items():
    records = parse_query_response(FIXTURE.read_text())
    r1 = records[1]
    assert len(r1.line_items) == 2
    assert r1.line_items[0].item_ref.full_name == "40-RAG12"
    assert r1.line_items[0].amount == Decimal("500.00")


def test_parse_empty_line_items_ok():
    records = parse_query_response(FIXTURE.read_text())
    assert records[0].line_items == []
