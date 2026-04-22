from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from qb_cli.models.invoice import Invoice
from qb_cli.ops.query_op import query
from tests.unit.ops._fakes import make_factory


_FIXTURE_DIR = Path(__file__).resolve().parents[2] / "fixtures" / "qbxml_responses"


def test_query_invoice_returns_parsed_entities():
    fixture = (_FIXTURE_DIR / "invoice_query_sample.xml").read_text()
    ctx = SimpleNamespace(connection_factory=make_factory(fixture))
    result = query(ctx, "invoice", ref_numbers=["14396", "14397"])
    assert len(result) == 2
    assert all(isinstance(r, Invoice) for r in result)
    assert {r.ref_number for r in result} == {"14396", "14397"}


def test_query_sales_order_returns_parsed_entities():
    fixture = (_FIXTURE_DIR / "sales_order_query_sample.xml").read_text()
    ctx = SimpleNamespace(connection_factory=make_factory(fixture))
    result = query(ctx, "sales_order")
    assert len(result) == 2


def test_query_unknown_entity_raises():
    ctx = SimpleNamespace(connection_factory=make_factory(""))
    with pytest.raises(KeyError):
        query(ctx, "bogus")
