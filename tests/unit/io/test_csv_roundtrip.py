from decimal import Decimal
from pathlib import Path
from qb_cli.models.invoice import Invoice, InvoiceLineItem
from qb_cli.models.sales_order import SalesOrder
from qb_cli.models.purchase_order import PurchaseOrder
from qb_cli.io.csv_serializer import to_csv, from_csv


def test_csv_handles_non_ascii_text(tmp_path: Path):
    """Regression: Windows default encoding (cp1252) can't represent smart
    quotes, em-dashes, etc., causing UnicodeEncodeError on writerow. The
    serializer must force UTF-8 on both write and read."""
    inv = Invoice(
        ref_number="14396",
        customer_ref={"FullName": "Café — Smith’s"},
        memo="Note “quoted” text",
        line_items=[{"item_ref": {"FullName": "40-RAG12"}, "quantity": "1", "amount": "1.00"}],
    )
    path = tmp_path / "unicode.csv"
    to_csv([inv], path)
    text = path.read_text(encoding="utf-8")
    assert "Café" in text
    loaded = from_csv(Invoice, path)
    assert loaded[0].customer_ref.full_name == "Café — Smith’s"
    assert loaded[0].memo == "Note “quoted” text"


def test_invoice_round_trip(tmp_path: Path):
    inv = Invoice(
        ref_number="14396",
        customer_ref={"FullName": "ACME Corp"},
        txn_date="2025-05-01",
        po_number="7740-SH",
        ship_address={"Addr1": "100 Main", "City": "Seattle", "State": "WA", "PostalCode": "98101"},
        line_items=[
            {"item_ref": {"FullName": "40-RAG12"}, "quantity": "10", "rate": "25.00", "amount": "250.00"},
            {"item_ref": {"FullName": "Freight"}, "quantity": "1", "rate": "35.00", "amount": "35.00"},
        ],
    )
    path = tmp_path / "inv.csv"
    to_csv([inv], path)

    text = path.read_text()
    assert "row_type" in text
    assert "parent_ref" in text

    loaded = from_csv(Invoice, path)
    assert len(loaded) == 1
    r = loaded[0]
    assert r.ref_number == "14396"
    assert r.customer_ref.full_name == "ACME Corp"
    assert r.po_number == "7740-SH"
    assert r.ship_address.city == "Seattle"
    assert len(r.line_items) == 2
    assert r.line_items[0].item_ref.full_name == "40-RAG12"
    assert r.line_items[0].amount == Decimal("250.00")


def test_sales_order_round_trip(tmp_path: Path):
    so = SalesOrder(
        ref_number="7740",
        customer_ref={"FullName": "ACME Corp"},
        line_items=[
            {"item_ref": {"FullName": "40-RAG12"}, "quantity": "10", "rate": "25.00", "amount": "250.00"},
        ],
    )
    path = tmp_path / "so.csv"
    to_csv([so], path)
    loaded = from_csv(SalesOrder, path)
    assert len(loaded) == 1
    assert loaded[0].line_items[0].amount == Decimal("250.00")


def test_purchase_order_round_trip(tmp_path: Path):
    po = PurchaseOrder(
        ref_number="7740",
        vendor_ref={"FullName": "Widgets Inc."},
        line_items=[
            {"item_ref": {"FullName": "40-RAG12"}, "quantity": "100", "rate": "5.00", "amount": "500.00"},
        ],
    )
    path = tmp_path / "po.csv"
    to_csv([po], path)
    loaded = from_csv(PurchaseOrder, path)
    assert loaded[0].vendor_ref.full_name == "Widgets Inc."
    assert loaded[0].line_items[0].amount == Decimal("500.00")


def test_multiple_invoices_with_mixed_line_counts(tmp_path: Path):
    inv1 = Invoice(
        ref_number="14396",
        customer_ref={"FullName": "A"},
        line_items=[{"item_ref": {"FullName": "X"}, "quantity": "1", "amount": "1.00"}],
    )
    inv2 = Invoice(
        ref_number="14397",
        customer_ref={"FullName": "B"},
        line_items=[
            {"item_ref": {"FullName": "Y"}, "quantity": "2", "amount": "2.00"},
            {"item_ref": {"FullName": "Z"}, "quantity": "3", "amount": "3.00"},
        ],
    )
    path = tmp_path / "multi.csv"
    to_csv([inv1, inv2], path)
    loaded = from_csv(Invoice, path)
    assert [r.ref_number for r in loaded] == ["14396", "14397"]
    assert len(loaded[0].line_items) == 1
    assert len(loaded[1].line_items) == 2


def test_invoice_without_line_items(tmp_path: Path):
    inv = Invoice(ref_number="14398", customer_ref={"FullName": "A"})
    path = tmp_path / "noline.csv"
    to_csv([inv], path)
    loaded = from_csv(Invoice, path)
    assert loaded[0].ref_number == "14398"
    assert loaded[0].line_items == []
