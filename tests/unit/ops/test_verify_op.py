from __future__ import annotations

from types import SimpleNamespace

from qb_cli.ops.verify_op import verify_entities


_CUSTOMER_RESPONSE_BOTH = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <CustomerQueryRs requestID="1" statusCode="0" statusSeverity="Info">
      <CustomerRet>
        <ListID>80000001-1</ListID>
        <Name>ACME</Name>
        <FullName>ACME</FullName>
      </CustomerRet>
      <CustomerRet>
        <ListID>80000001-2</ListID>
        <Name>Other Co</Name>
        <FullName>Other Co</FullName>
      </CustomerRet>
    </CustomerQueryRs>
  </QBXMLMsgsRs>
</QBXML>
"""

_CUSTOMER_RESPONSE_PARTIAL = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <CustomerQueryRs requestID="1" statusCode="0" statusSeverity="Info">
      <CustomerRet>
        <ListID>80000001-1</ListID>
        <Name>ACME</Name>
        <FullName>ACME</FullName>
      </CustomerRet>
    </CustomerQueryRs>
  </QBXMLMsgsRs>
</QBXML>
"""

_CUSTOMER_RESPONSE_NONE = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <CustomerQueryRs requestID="1" statusCode="500" statusSeverity="Info" statusMessage="A query request did not find a matching record."/>
  </QBXMLMsgsRs>
</QBXML>
"""


def _factory_returning(response: str):
    class FakeConn:
        def __enter__(self_inner):
            return self_inner

        def __exit__(self_inner, *a):
            return False

        def send(self_inner, req):
            return response

    def make():
        return FakeConn()

    return make


def test_verify_all_exist():
    ctx = SimpleNamespace(connection_factory=_factory_returning(_CUSTOMER_RESPONSE_BOTH))
    result = verify_entities(ctx, customers=["ACME", "Other Co"])
    assert result.found["customer"] == {"ACME", "Other Co"}
    assert result.missing["customer"] == set()
    assert result.all_found


def test_verify_none_exist_status_500():
    ctx = SimpleNamespace(connection_factory=_factory_returning(_CUSTOMER_RESPONSE_NONE))
    result = verify_entities(ctx, customers=["MissingCo"])
    assert result.found["customer"] == set()
    assert result.missing["customer"] == {"MissingCo"}
    assert not result.all_found


def test_verify_partial():
    ctx = SimpleNamespace(connection_factory=_factory_returning(_CUSTOMER_RESPONSE_PARTIAL))
    result = verify_entities(ctx, customers=["ACME", "Nope1", "Nope2"])
    assert result.found["customer"] == {"ACME"}
    assert result.missing["customer"] == {"Nope1", "Nope2"}


def test_verify_empty_inputs_returns_empty_buckets():
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
    result = verify_entities(ctx)
    assert result.found == {"customer": set(), "vendor": set(), "item": set(), "terms": set()}
    assert result.missing == {"customer": set(), "vendor": set(), "item": set(), "terms": set()}
    assert called["n"] == 0


def test_emitted_qbxml_uses_fullname_for_every_bucket():
    """Regression: TermsQueryRq previously used <Name>, which fails QB schema validation.
    Also regression: IncludeRetElement on Customer/Vendor triggered parser errors on some
    QB versions. Pin the emitted shape so neither comes back.
    """
    sent: list[str] = []

    class FakeConn:
        def __enter__(self_inner):
            return self_inner

        def __exit__(self_inner, *a):
            return False

        def send(self_inner, req):
            sent.append(req)
            return _CUSTOMER_RESPONSE_NONE  # 500 — treated as all-missing

    ctx = SimpleNamespace(connection_factory=lambda: FakeConn())
    verify_entities(
        ctx,
        customers=["C1"],
        vendors=["V1"],
        items=["I1"],
        terms=["T1"],
    )
    assert len(sent) == 4
    for req in sent:
        assert "<FullName>" in req, f"missing <FullName> in: {req[:200]}"
        assert "IncludeRetElement" not in req, "IncludeRetElement causes parser errors on some QB versions"
    terms_req = next(r for r in sent if "TermsQueryRq" in r)
    assert "<FullName>T1</FullName>" in terms_req
    assert "<Name>T1</Name>" not in terms_req


def test_verify_multiple_buckets_dispatches_correctly():
    vendor_response = """<?xml version="1.0" ?>
<QBXML>
  <QBXMLMsgsRs>
    <VendorQueryRs requestID="1" statusCode="0" statusSeverity="Info">
      <VendorRet>
        <ListID>v-1</ListID>
        <Name>Widgets Inc</Name>
        <FullName>Widgets Inc</FullName>
      </VendorRet>
    </VendorQueryRs>
  </QBXMLMsgsRs>
</QBXML>
"""

    class FakeConn:
        def __enter__(self_inner):
            return self_inner

        def __exit__(self_inner, *a):
            return False

        def send(self_inner, req):
            if "VendorQueryRq" in req:
                return vendor_response
            if "CustomerQueryRq" in req:
                return _CUSTOMER_RESPONSE_PARTIAL
            raise AssertionError(f"unexpected request: {req[:200]}")

    ctx = SimpleNamespace(connection_factory=lambda: FakeConn())
    result = verify_entities(ctx, customers=["ACME", "Bogus"], vendors=["Widgets Inc"])
    assert result.found["customer"] == {"ACME"}
    assert result.missing["customer"] == {"Bogus"}
    assert result.found["vendor"] == {"Widgets Inc"}
    assert result.missing["vendor"] == set()
