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
