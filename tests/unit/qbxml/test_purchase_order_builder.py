from datetime import date
from decimal import Decimal
import pytest
from qb_cli.models.purchase_order import PurchaseOrder, PurchaseOrderLineItem
from qb_cli.models.shared import Ref
from qb_cli.qbxml.purchase_order import build_query, build_add, build_mod


def test_build_query_by_ref_numbers():
    out = build_query(ref_numbers=["7740", "7741"])
    assert '<RefNumber>7740</RefNumber>' in out
    assert '<RefNumber>7741</RefNumber>' in out
    assert 'PurchaseOrderQueryRq' in out


def test_build_query_date_range_order_valid():
    out = build_query(date_from=date(2025, 1, 1), date_to=date(2025, 12, 31))
    idx_filter = out.index('TxnDateRangeFilter')
    idx_include = out.index('IncludeLineItems')
    assert idx_filter < idx_include


def test_build_query_with_ref_does_not_emit_max_returned():
    out = build_query(ref_numbers=["7740"], max_returned=100)
    assert 'MaxReturned' not in out


def test_build_add_emits_vendor_and_ref():
    po = PurchaseOrder(
        vendor_ref=Ref(full_name="Widgets Inc."),
        ref_number="7740",
        memo="Test & demo",
        line_items=[PurchaseOrderLineItem(item_ref=Ref(full_name="40-RAG12"), quantity=Decimal("100"))],
    )
    out = build_add(po)
    assert '<VendorRef>' in out
    assert '<FullName>Widgets Inc.</FullName>' in out
    assert '<RefNumber>7740</RefNumber>' in out
    assert '<PurchaseOrderLineAdd>' in out
    assert '<Memo>Test &amp; demo</Memo>' in out


def test_build_mod_requires_edit_sequence():
    po = PurchaseOrder(txn_id="PO-1", vendor_ref=Ref(full_name="X"))
    with pytest.raises(ValueError):
        build_mod(po)


def test_build_mod_emits_txn_id_and_edit_sequence():
    po = PurchaseOrder(txn_id="PO-1", edit_sequence="42", vendor_ref=Ref(full_name="X"))
    out = build_mod(po)
    assert '<TxnID>PO-1</TxnID>' in out
    assert '<EditSequence>42</EditSequence>' in out
