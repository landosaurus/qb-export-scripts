from datetime import date
from decimal import Decimal
import pytest
from qb_cli.models.invoice import Invoice, InvoiceLineItem
from qb_cli.models.shared import Ref
from qb_cli.qbxml.invoice import build_query, build_add, build_mod


def test_build_query_by_ref_numbers():
    out = build_query(ref_numbers=["14396", "14397"], include_line_items=True)
    assert '<RefNumber>14396</RefNumber>' in out
    assert '<RefNumber>14397</RefNumber>' in out
    assert '<IncludeLineItems>true</IncludeLineItems>' in out
    assert '<?qbxml version="16.0"?>' in out


def test_build_query_date_range_order_valid():
    out = build_query(date_from=date(2025, 1, 1), date_to=date(2025, 12, 31))
    idx_filter = out.index('TxnDateRangeFilter')
    idx_include = out.index('IncludeLineItems')
    assert idx_filter < idx_include


def test_build_query_with_ref_does_not_emit_max_returned():
    out = build_query(ref_numbers=["14396"], max_returned=100)
    assert 'MaxReturned' not in out


def test_build_add_emits_customer_and_ref():
    inv = Invoice(
        customer_ref=Ref(full_name="ACME Corp"),
        ref_number="14396",
        memo="Test & demo",
        line_items=[InvoiceLineItem(item_ref=Ref(full_name="40-RAG12"), quantity=Decimal("10"), rate=Decimal("25.00"))],
    )
    out = build_add(inv)
    assert '<CustomerRef>' in out
    assert '<FullName>ACME Corp</FullName>' in out
    assert '<RefNumber>14396</RefNumber>' in out
    assert '<InvoiceLineAdd>' in out
    # Memo ampersand must be escaped
    assert '&amp;' in out
    assert '<Memo>Test &amp; demo</Memo>' in out


def test_build_mod_requires_edit_sequence():
    inv = Invoice(txn_id="ABC-1", customer_ref=Ref(full_name="X"))
    with pytest.raises(ValueError):
        build_mod(inv)


def test_build_mod_emits_txn_id_and_edit_sequence():
    inv = Invoice(txn_id="ABC-1", edit_sequence="42", customer_ref=Ref(full_name="X"))
    out = build_mod(inv)
    assert '<TxnID>ABC-1</TxnID>' in out
    assert '<EditSequence>42</EditSequence>' in out
