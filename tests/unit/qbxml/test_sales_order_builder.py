from datetime import date
from decimal import Decimal
import pytest
from qb_cli.models.sales_order import SalesOrder, SalesOrderLineItem
from qb_cli.models.shared import Ref
from qb_cli.qbxml.sales_order import build_query, build_add, build_mod


def test_build_query_by_ref_numbers():
    out = build_query(ref_numbers=["7740", "7741"])
    assert '<RefNumber>7740</RefNumber>' in out
    assert '<RefNumber>7741</RefNumber>' in out
    assert 'SalesOrderQueryRq' in out


def test_build_query_date_range_order_valid():
    out = build_query(date_from=date(2025, 1, 1), date_to=date(2025, 12, 31))
    idx_filter = out.index('TxnDateRangeFilter')
    idx_include = out.index('IncludeLineItems')
    assert idx_filter < idx_include


def test_build_query_with_ref_does_not_emit_max_returned():
    out = build_query(ref_numbers=["7740"], max_returned=100)
    assert 'MaxReturned' not in out


def test_build_add_emits_customer_and_ref():
    so = SalesOrder(
        customer_ref=Ref(full_name="ACME Corp"),
        ref_number="7740",
        memo="Test & demo",
        line_items=[SalesOrderLineItem(item_ref=Ref(full_name="40-RAG12"), quantity=Decimal("10"))],
    )
    out = build_add(so)
    assert '<CustomerRef>' in out
    assert '<FullName>ACME Corp</FullName>' in out
    assert '<SalesOrderLineAdd>' in out
    assert '<Memo>Test &amp; demo</Memo>' in out


def test_build_mod_requires_edit_sequence():
    so = SalesOrder(txn_id="SO-1", customer_ref=Ref(full_name="X"))
    with pytest.raises(ValueError):
        build_mod(so)


def test_build_mod_emits_txn_id_and_edit_sequence():
    so = SalesOrder(txn_id="SO-1", edit_sequence="42", customer_ref=Ref(full_name="X"))
    out = build_mod(so)
    assert '<TxnID>SO-1</TxnID>' in out
    assert '<EditSequence>42</EditSequence>' in out
