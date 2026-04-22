from qb_cli.qbxml.envelope import wrap_request


def test_wrap_adds_qbxml_header_and_msgs():
    inner = '<InvoiceQueryRq requestID="1"></InvoiceQueryRq>'
    out = wrap_request(inner)
    assert out.startswith('<?xml')
    assert '<?qbxml version="16.0"?>' in out
    assert '<QBXMLMsgsRq onError="continueOnError">' in out
    assert inner in out
    assert out.strip().endswith("</QBXML>")
