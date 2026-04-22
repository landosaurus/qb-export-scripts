from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from qb_cli.ops.dedupe import find_duplicates
from tests.unit.ops._fakes import make_factory


_FIXTURE_DIR = Path(__file__).resolve().parents[2] / "fixtures" / "qbxml_responses"


_EMPTY_500 = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <InvoiceQueryRs requestID="1" statusCode="500" statusSeverity="Info" statusMessage="No matches"/>
  </QBXMLMsgsRs>
</QBXML>
"""


def test_find_duplicates_all_exist():
    fixture = (_FIXTURE_DIR / "invoice_query_sample.xml").read_text()
    ctx = SimpleNamespace(connection_factory=make_factory(fixture))
    result = find_duplicates(ctx, "invoice", ["14396", "14397"])
    assert result == {"14396", "14397"}


def test_find_duplicates_partial():
    fixture = (_FIXTURE_DIR / "invoice_query_sample.xml").read_text()
    ctx = SimpleNamespace(connection_factory=make_factory(fixture))
    result = find_duplicates(ctx, "invoice", ["14396", "14397", "99999"])
    assert result == {"14396", "14397"}


def test_find_duplicates_none_status_500():
    ctx = SimpleNamespace(connection_factory=make_factory(_EMPTY_500))
    result = find_duplicates(ctx, "invoice", ["NOPE1", "NOPE2"])
    assert result == set()


def test_find_duplicates_empty_input():
    called = {"n": 0}

    class FakeConn:
        def __enter__(self_inner):
            called["n"] += 1
            return self_inner

        def __exit__(self_inner, *a):
            return False

        def send(self_inner, req):
            raise AssertionError("should not be called")

    ctx = SimpleNamespace(connection_factory=lambda: FakeConn())
    result = find_duplicates(ctx, "invoice", [])
    assert result == set()
    assert called["n"] == 0


def test_find_duplicates_purchase_order():
    fixture = (_FIXTURE_DIR / "purchase_order_query_sample.xml").read_text()
    ctx = SimpleNamespace(connection_factory=make_factory(fixture))
    result = find_duplicates(ctx, "purchase_order", ["7740", "7741", "unknown"])
    assert result == {"7740", "7741"}
