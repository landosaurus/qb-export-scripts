import pytest
from decimal import Decimal
from pydantic import ValidationError
from qb_cli.models.sales_order import SalesOrder, SalesOrderLineItem
from qb_cli.models.shared import Ref


def test_minimal_required_fields_accepted():
    so = SalesOrder(customer_ref=Ref(full_name="ACME Corp"), ref_number="7740")
    assert so.ref_number == "7740"


def test_ref_number_max_length():
    with pytest.raises(ValidationError):
        SalesOrder(customer_ref=Ref(full_name="X"), ref_number="X" * 12)


def test_qb_alias_round_trip():
    payload = {
        "RefNumber": "7740",
        "CustomerRef": {"FullName": "ACME Corp"},
        "TxnDate": "2025-05-01",
        "ShipDate": "2025-05-15",
        "IsManuallyClosed": False,
    }
    so = SalesOrder.model_validate(payload)
    dumped = so.model_dump(by_alias=True, exclude_none=True, mode="json")
    assert dumped["RefNumber"] == "7740"
    assert dumped["IsManuallyClosed"] is False


def test_populates_line_items_from_qb_casing():
    payload = {
        "RefNumber": "7740",
        "CustomerRef": {"FullName": "ACME Corp"},
        "SalesOrderLineRet": [
            {
                "ItemRef": {"FullName": "40-RAG12"},
                "Quantity": "10",
                "Rate": "25.00",
                "Amount": "250.00",
            }
        ],
    }
    so = SalesOrder.model_validate(payload)
    assert len(so.line_items) == 1
    assert so.line_items[0].amount == Decimal("250.00")


def test_required_on_add_is_enforced_at_class_level():
    assert "customer_ref" in SalesOrder.REQUIRED_ON_ADD
