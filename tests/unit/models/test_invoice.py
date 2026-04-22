import pytest
from decimal import Decimal
from datetime import date
from pydantic import ValidationError
from qb_cli.models.invoice import Invoice, InvoiceLineItem
from qb_cli.models.shared import Address, Ref


def test_minimal_required_fields_accepted() -> None:
    inv = Invoice(customer_ref=Ref(full_name="ACME Corp"), ref_number="14396")
    assert inv.ref_number == "14396"
    assert inv.customer_ref.full_name == "ACME Corp"


def test_ref_number_max_length() -> None:
    with pytest.raises(ValidationError):
        Invoice(customer_ref=Ref(full_name="X"), ref_number="X" * 12)


def test_qb_alias_round_trip() -> None:
    payload = {
        "RefNumber": "14396",
        "CustomerRef": {"FullName": "ACME Corp"},
        "TxnDate": "2025-05-01",
        "PONumber": "7740-SH",
        "ShipAddress": {"Addr1": "100 Main", "City": "Seattle", "State": "WA", "PostalCode": "98101"},
    }
    inv = Invoice.model_validate(payload)
    dumped = inv.model_dump(by_alias=True, exclude_none=True, mode="json")
    assert dumped["RefNumber"] == "14396"
    assert dumped["CustomerRef"] == {"FullName": "ACME Corp"}
    assert dumped["TxnDate"] == "2025-05-01"
    assert dumped["ShipAddress"]["Addr1"] == "100 Main"


def test_populates_line_items_from_qb_casing() -> None:
    payload = {
        "RefNumber": "14396",
        "CustomerRef": {"FullName": "ACME Corp"},
        "InvoiceLineRet": [
            {
                "ItemRef": {"FullName": "40-RAG12"},
                "Desc": "Reclaimed terry",
                "Quantity": "10",
                "Rate": "25.00",
                "Amount": "250.00",
            }
        ],
    }
    inv = Invoice.model_validate(payload)
    assert len(inv.line_items) == 1
    assert inv.line_items[0].item_ref.full_name == "40-RAG12"
    assert inv.line_items[0].amount == Decimal("250.00")


def test_required_on_add_is_enforced_at_class_level() -> None:
    from qb_cli.models.invoice import Invoice
    assert "customer_ref" in Invoice.REQUIRED_ON_ADD


def test_memo_max_length() -> None:
    with pytest.raises(ValidationError):
        Invoice(customer_ref=Ref(full_name="X"), memo="m" * 4096)
