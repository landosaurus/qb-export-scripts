from __future__ import annotations

from datetime import date
from typing import Optional

from lxml import etree

from qb_cli.models.shared import Address


def xml_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
         .replace("<", "&lt;")
         .replace(">", "&gt;")
         .replace("'", "&apos;")
         .replace('"', "&quot;")
    )


def format_qb_date(d: date) -> str:
    return d.isoformat()


_ADDR_TAGS = ("Addr1", "Addr2", "Addr3", "Addr4", "Addr5",
              "City", "State", "PostalCode", "Country", "Note")


def parse_address(elem: Optional[etree._Element]) -> Optional[Address]:
    if elem is None:
        return None
    payload: dict[str, str] = {}
    for tag in _ADDR_TAGS:
        child = elem.find(tag)
        if child is not None and child.text:
            payload[tag] = child.text.strip()
    if not payload:
        return None
    return Address.model_validate(payload)


def read_iterator_state(response_xml: str) -> tuple[Optional[str], int]:
    """Extract iterator state from a QBXML response.

    Returns ``(iterator_id, remaining_count)``:
    - ``iterator_id`` is the ``iteratorID`` attribute on the first ``*QueryRs`` element,
      or ``None`` if the response is not iterated.
    - ``remaining_count`` is the ``iteratorRemainingCount`` attribute as int, or 0
      if missing (meaning there are no more pages).
    """
    root = etree.fromstring(response_xml.encode("utf-8"))
    for rs in root.xpath("//*[local-name()='QBXMLMsgsRs']/*"):
        iid = rs.get("iteratorID")
        remaining = int(rs.get("iteratorRemainingCount", "0"))
        return iid, remaining
    return None, 0
