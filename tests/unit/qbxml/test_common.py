from datetime import date
from qb_cli.qbxml.common import xml_escape, parse_address, format_qb_date


def test_xml_escape():
    assert xml_escape("A & B < C > 'x' \"y\"") == "A &amp; B &lt; C &gt; &apos;x&apos; &quot;y&quot;"


def test_format_qb_date():
    assert format_qb_date(date(2025, 5, 1)) == "2025-05-01"


def test_parse_address_from_element():
    from lxml import etree
    xml = b"""<ShipAddress>
        <Addr1>100 Main</Addr1><City>Seattle</City><State>WA</State>
        <PostalCode>98101</PostalCode>
    </ShipAddress>"""
    elem = etree.fromstring(xml)
    addr = parse_address(elem)
    assert addr.addr1 == "100 Main"
    assert addr.city == "Seattle"
