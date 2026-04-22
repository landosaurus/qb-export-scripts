from __future__ import annotations

import pytest

from qb_cli.models.invoice import Invoice
from qb_cli.models.purchase_order import PurchaseOrder
from qb_cli.models.sales_order import SalesOrder
from qb_cli.ops.registry import EntityHandler, get_handler


def test_invoice_handler():
    h = get_handler("invoice")
    assert isinstance(h, EntityHandler)
    assert h.key == "invoice"
    assert h.model is Invoice
    assert h.id_field == "ref_number"
    assert hasattr(h.qbxml, "build_query")
    assert hasattr(h.qbxml, "parse_query_response")


def test_sales_order_handler():
    h = get_handler("sales_order")
    assert h.model is SalesOrder
    assert h.key == "sales_order"


def test_purchase_order_handler():
    h = get_handler("purchase_order")
    assert h.model is PurchaseOrder
    assert h.key == "purchase_order"


def test_unknown_key_raises_keyerror_with_supported_keys():
    with pytest.raises(KeyError) as excinfo:
        get_handler("widget")
    msg = str(excinfo.value)
    assert "widget" in msg
    assert "invoice" in msg
    assert "sales_order" in msg
    assert "purchase_order" in msg
