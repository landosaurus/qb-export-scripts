from decimal import Decimal

from qb_cli.models.shared import Address, Ref, quantize_money


def test_address_render_multiline() -> None:
    a = Address(addr1="100 Main", city="Seattle", state="WA", postal_code="98101")
    assert a.render_multiline() == "100 Main\nSeattle, WA 98101"


def test_address_accepts_qb_casing() -> None:
    a = Address.model_validate({"Addr1": "100 Main", "City": "X", "State": "Y", "PostalCode": "1"})
    assert a.addr1 == "100 Main"


def test_ref_requires_full_name_or_list_id() -> None:
    r = Ref(full_name="ACME Corp")
    assert r.full_name == "ACME Corp"


def test_money_quantization() -> None:
    assert quantize_money(Decimal("10.1")) == Decimal("10.10")
    assert quantize_money(Decimal("10.123")) == Decimal("10.12")


def test_money_annotated_type_auto_quantizes_on_amount_fields() -> None:
    """Regression for QB 3040: Excel strips trailing zeros from exported CSVs
    (97.50 -> 97.5), and QuickBooks rejects any Amount without exactly 2dp.
    The Money annotated type ensures every amount is stored as 2dp regardless
    of input precision. Rate/Quantity are NOT money and keep their precision."""
    from qb_cli.models.sales_order import SalesOrder, SalesOrderLineItem

    line = SalesOrderLineItem.model_validate({
        "Amount": "97.5",
        "Rate": "3.9",
        "Quantity": "25",
    })
    assert line.amount == Decimal("97.50")
    assert str(line.amount) == "97.50"
    assert line.rate == Decimal("3.9")
    assert line.quantity == Decimal("25")

    so = SalesOrder.model_validate({
        "RefNumber": "7959",
        "CustomerRef": {"FullName": "ACME"},
        "Subtotal": "156.75",
        "TotalAmount": "156.75",
    })
    assert str(so.subtotal) == "156.75"
    assert str(so.total_amount) == "156.75"

    line2 = SalesOrderLineItem.model_validate({"Amount": 100})
    assert str(line2.amount) == "100.00"

    line3 = SalesOrderLineItem.model_validate({})
    assert line3.amount is None
