import pytest
from decimal import Decimal
from pydantic import ValidationError
from qb_cli.models.purchase_order import PurchaseOrder, PurchaseOrderLineItem
from qb_cli.models.shared import Ref


def test_minimal_required_fields_accepted():
    po = PurchaseOrder(vendor_ref=Ref(full_name="SomeVendor"), ref_number="7740")
    assert po.vendor_ref.full_name == "SomeVendor"
    assert po.ref_number == "7740"


def test_ref_number_max_length():
    with pytest.raises(ValidationError):
        PurchaseOrder(vendor_ref=Ref(full_name="X"), ref_number="X" * 12)


def test_qb_alias_round_trip():
    payload = {
        "RefNumber": "7740",
        "VendorRef": {"FullName": "SomeVendor"},
        "TxnDate": "2025-05-01",
        "ExpectedDate": "2025-05-20",
        "VendorAddress": {"Addr1": "1 Mfg Rd", "City": "Portland", "State": "OR", "PostalCode": "97201"},
    }
    po = PurchaseOrder.model_validate(payload)
    dumped = po.model_dump(by_alias=True, exclude_none=True, mode="json")
    assert dumped["RefNumber"] == "7740"
    assert dumped["VendorRef"] == {"FullName": "SomeVendor"}
    assert dumped["ExpectedDate"] == "2025-05-20"
    assert dumped["VendorAddress"]["Addr1"] == "1 Mfg Rd"


def test_populates_line_items_from_qb_casing():
    payload = {
        "RefNumber": "7740",
        "VendorRef": {"FullName": "SomeVendor"},
        "PurchaseOrderLineRet": [
            {
                "ItemRef": {"FullName": "40-RAG12"},
                "Quantity": "100",
                "Rate": "5.00",
                "Amount": "500.00",
            }
        ],
    }
    po = PurchaseOrder.model_validate(payload)
    assert len(po.line_items) == 1
    assert po.line_items[0].amount == Decimal("500.00")


def test_required_on_add_is_enforced_at_class_level():
    assert "vendor_ref" in PurchaseOrder.REQUIRED_ON_ADD
