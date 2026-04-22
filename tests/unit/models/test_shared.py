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
