from decimal import Decimal
from pathlib import Path
from qb_cli.models.invoice import Invoice
from qb_cli.models.sales_order import SalesOrder
from qb_cli.io.json_serializer import to_json, from_json


def test_invoice_round_trip(tmp_path: Path):
    original = Invoice(
        ref_number="14396",
        customer_ref={"FullName": "ACME Corp"},
        txn_date="2025-05-01",
        line_items=[
            {"item_ref": {"FullName": "40-RAG12"}, "quantity": "10", "rate": "25.00", "amount": "250.00"},
        ],
    )
    path = tmp_path / "inv.json"
    to_json([original], path)
    loaded = from_json(Invoice, path)
    assert len(loaded) == 1
    assert loaded[0].ref_number == original.ref_number
    assert loaded[0].line_items[0].amount == Decimal("250.00")


def test_sales_order_round_trip(tmp_path: Path):
    original = SalesOrder(
        ref_number="7740",
        customer_ref={"FullName": "ACME Corp"},
        line_items=[
            {"item_ref": {"FullName": "40-RAG12"}, "quantity": "10", "rate": "25.00", "amount": "250.00"},
        ],
    )
    path = tmp_path / "so.json"
    to_json([original], path)
    loaded = from_json(SalesOrder, path)
    assert loaded[0].line_items[0].amount == Decimal("250.00")


def test_empty_list_round_trip(tmp_path: Path):
    path = tmp_path / "empty.json"
    to_json([], path)
    loaded = from_json(Invoice, path)
    assert loaded == []
