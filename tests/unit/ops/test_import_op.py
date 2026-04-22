from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace
from typing import Callable, List

from qb_cli.ops.import_op import import_


_VERIFY_CUSTOMER_OK = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <CustomerQueryRs requestID="1" statusCode="0" statusSeverity="Info">
      <CustomerRet>
        <ListID>c-1</ListID>
        <Name>ACME Corp</Name>
        <FullName>ACME Corp</FullName>
      </CustomerRet>
    </CustomerQueryRs>
  </QBXMLMsgsRs>
</QBXML>
"""

_VERIFY_CUSTOMER_MISSING = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <CustomerQueryRs requestID="1" statusCode="500" statusSeverity="Info" statusMessage="No match"/>
  </QBXMLMsgsRs>
</QBXML>
"""

_VERIFY_ITEM_OK = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <ItemQueryRs requestID="1" statusCode="0" statusSeverity="Info">
      <ItemServiceRet>
        <ListID>i-1</ListID>
        <Name>Widget</Name>
        <FullName>Widget</FullName>
      </ItemServiceRet>
    </ItemQueryRs>
  </QBXMLMsgsRs>
</QBXML>
"""

_DEDUPE_NONE = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <InvoiceQueryRs requestID="1" statusCode="500" statusSeverity="Info" statusMessage="No matches"/>
  </QBXMLMsgsRs>
</QBXML>
"""

_DEDUPE_ONE_EXISTS = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <InvoiceQueryRs requestID="1" statusCode="0" statusSeverity="Info">
      <InvoiceRet>
        <TxnID>ABC-1</TxnID>
        <EditSequence>99</EditSequence>
        <TxnNumber>1</TxnNumber>
        <CustomerRef><FullName>ACME Corp</FullName></CustomerRef>
        <TxnDate>2025-05-01</TxnDate>
        <RefNumber>INV-001</RefNumber>
      </InvoiceRet>
    </InvoiceQueryRs>
  </QBXMLMsgsRs>
</QBXML>
"""

_INVOICE_ADD_OK = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <InvoiceAddRs requestID="1" statusCode="0" statusSeverity="Info">
      <InvoiceRet>
        <TxnID>new-txn</TxnID>
        <EditSequence>1</EditSequence>
        <TxnNumber>9001</TxnNumber>
        <CustomerRef><FullName>ACME Corp</FullName></CustomerRef>
        <TxnDate>2025-05-01</TxnDate>
        <RefNumber>INV-001</RefNumber>
      </InvoiceRet>
    </InvoiceAddRs>
  </QBXMLMsgsRs>
</QBXML>
"""

_INVOICE_MOD_OK = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <InvoiceModRs requestID="1" statusCode="0" statusSeverity="Info">
      <InvoiceRet>
        <TxnID>ABC-1</TxnID>
        <EditSequence>100</EditSequence>
        <TxnNumber>1</TxnNumber>
        <CustomerRef><FullName>ACME Corp</FullName></CustomerRef>
        <TxnDate>2025-05-01</TxnDate>
        <RefNumber>INV-001</RefNumber>
      </InvoiceRet>
    </InvoiceModRs>
  </QBXMLMsgsRs>
</QBXML>
"""


def _write_invoice_json(path: Path, ref_numbers: list[str]) -> None:
    records = []
    for rn in ref_numbers:
        records.append(
            {
                "CustomerRef": {"FullName": "ACME Corp"},
                "TxnDate": "2025-05-01",
                "RefNumber": rn,
                "InvoiceLineRet": [
                    {"ItemRef": {"FullName": "Widget"}, "Quantity": "1"}
                ],
            }
        )
    path.write_text(json.dumps(records))


class _RouterConn:
    """Fake connection that dispatches by request-type markers in the XML."""

    def __init__(self, router: Callable[[str], str], sent_requests: List[str]) -> None:
        self._router = router
        self._sent = sent_requests

    def __enter__(self) -> "_RouterConn":
        return self

    def __exit__(self, *a: object) -> None:
        return None

    def send(self, req: str) -> str:
        self._sent.append(req)
        return self._router(req)


def _factory(router: Callable[[str], str]) -> tuple[Callable[[], _RouterConn], List[str]]:
    sent: List[str] = []

    def make() -> _RouterConn:
        return _RouterConn(router, sent)

    return make, sent


def test_dry_run_missing_customer_does_not_send_add(tmp_path: Path):
    path = tmp_path / "inv.json"
    _write_invoice_json(path, ["INV-001"])

    def route(req: str) -> str:
        if "CustomerQueryRq" in req:
            return _VERIFY_CUSTOMER_MISSING
        if "ItemQueryRq" in req:
            return _VERIFY_ITEM_OK
        raise AssertionError(f"unexpected: {req[:200]}")

    factory, sent = _factory(route)
    ctx = SimpleNamespace(connection_factory=factory)

    result = import_(ctx, "invoice", input_path=path, dry_run=True)
    assert result.written == 0
    assert any(s.startswith("customer:") for s in result.missing_refs)
    assert not any("InvoiceAddRq" in r for r in sent)
    assert not any("InvoiceModRq" in r for r in sent)


def test_happy_path_two_invoices_added(tmp_path: Path):
    path = tmp_path / "inv.json"
    _write_invoice_json(path, ["INV-001", "INV-002"])

    def route(req: str) -> str:
        if "CustomerQueryRq" in req:
            return _VERIFY_CUSTOMER_OK
        if "ItemQueryRq" in req:
            return _VERIFY_ITEM_OK
        if "InvoiceQueryRq" in req:
            return _DEDUPE_NONE
        if "InvoiceAddRq" in req:
            return _INVOICE_ADD_OK
        raise AssertionError(f"unexpected: {req[:200]}")

    factory, sent = _factory(route)
    ctx = SimpleNamespace(connection_factory=factory)

    result = import_(ctx, "invoice", input_path=path)
    assert result.written == 2
    assert result.failed == 0
    assert result.attempted == 2
    assert result.missing_refs == []
    assert result.duplicate_refs == []
    assert sum(1 for r in sent if "InvoiceAddRq" in r) == 2


def test_on_duplicate_error_aborts(tmp_path: Path):
    path = tmp_path / "inv.json"
    _write_invoice_json(path, ["INV-001"])

    def route(req: str) -> str:
        if "CustomerQueryRq" in req:
            return _VERIFY_CUSTOMER_OK
        if "ItemQueryRq" in req:
            return _VERIFY_ITEM_OK
        if "InvoiceQueryRq" in req:
            return _DEDUPE_ONE_EXISTS
        raise AssertionError(f"unexpected: {req[:200]}")

    factory, sent = _factory(route)
    ctx = SimpleNamespace(connection_factory=factory)

    result = import_(ctx, "invoice", input_path=path, on_duplicate="error")
    assert result.written == 0
    assert result.duplicate_refs == ["INV-001"]
    assert not any("InvoiceAddRq" in r for r in sent)


def test_on_duplicate_skip(tmp_path: Path):
    path = tmp_path / "inv.json"
    _write_invoice_json(path, ["INV-001", "INV-002"])

    def route(req: str) -> str:
        if "CustomerQueryRq" in req:
            return _VERIFY_CUSTOMER_OK
        if "ItemQueryRq" in req:
            return _VERIFY_ITEM_OK
        if "InvoiceQueryRq" in req:
            return _DEDUPE_ONE_EXISTS
        if "InvoiceAddRq" in req:
            return _INVOICE_ADD_OK
        raise AssertionError(f"unexpected: {req[:200]}")

    factory, sent = _factory(route)
    ctx = SimpleNamespace(connection_factory=factory)

    result = import_(ctx, "invoice", input_path=path, on_duplicate="skip")
    assert result.written == 1
    assert result.skipped == 1
    assert result.duplicate_refs == ["INV-001"]
    assert sum(1 for r in sent if "InvoiceAddRq" in r) == 1


def test_on_duplicate_update_sends_mod(tmp_path: Path):
    path = tmp_path / "inv.json"
    _write_invoice_json(path, ["INV-001"])

    def route(req: str) -> str:
        if "CustomerQueryRq" in req:
            return _VERIFY_CUSTOMER_OK
        if "ItemQueryRq" in req:
            return _VERIFY_ITEM_OK
        if "InvoiceQueryRq" in req:
            return _DEDUPE_ONE_EXISTS
        if "InvoiceModRq" in req:
            return _INVOICE_MOD_OK
        raise AssertionError(f"unexpected: {req[:200]}")

    factory, sent = _factory(route)
    ctx = SimpleNamespace(connection_factory=factory)

    result = import_(ctx, "invoice", input_path=path, on_duplicate="update")
    assert result.written == 1
    assert result.failed == 0
    assert result.duplicate_refs == ["INV-001"]
    mod_sends = [r for r in sent if "InvoiceModRq" in r]
    assert len(mod_sends) == 1
    assert "<EditSequence>99</EditSequence>" in mod_sends[0]
    assert "<TxnID>ABC-1</TxnID>" in mod_sends[0]
