from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Sequence

from lxml import etree

from qb_cli.context import Context
from qb_cli.qbxml.common import xml_escape
from qb_cli.qbxml.envelope import wrap_request
from qb_cli.transport.errors import QBEntityNotFound
from qb_cli.transport.status import check_response_status


@dataclass
class VerifyResult:
    found: dict[str, set[str]] = field(default_factory=dict)
    missing: dict[str, set[str]] = field(default_factory=dict)

    @property
    def all_found(self) -> bool:
        return all(len(v) == 0 for v in self.missing.values())


def _build_customer_query(names: Sequence[str]) -> str:
    parts: list[str] = ['    <CustomerQueryRq requestID="1">']
    for name in names:
        parts.append(f"      <FullName>{xml_escape(name)}</FullName>")
    parts.append("      <IncludeRetElement>Name</IncludeRetElement>")
    parts.append("      <IncludeRetElement>FullName</IncludeRetElement>")
    parts.append("      <IncludeRetElement>ListID</IncludeRetElement>")
    parts.append("    </CustomerQueryRq>")
    return wrap_request("\n".join(parts))


def _build_vendor_query(names: Sequence[str]) -> str:
    parts: list[str] = ['    <VendorQueryRq requestID="1">']
    for name in names:
        parts.append(f"      <FullName>{xml_escape(name)}</FullName>")
    parts.append("      <IncludeRetElement>Name</IncludeRetElement>")
    parts.append("      <IncludeRetElement>FullName</IncludeRetElement>")
    parts.append("      <IncludeRetElement>ListID</IncludeRetElement>")
    parts.append("    </VendorQueryRq>")
    return wrap_request("\n".join(parts))


def _build_item_query(names: Sequence[str]) -> str:
    parts: list[str] = ['    <ItemQueryRq requestID="1">']
    for name in names:
        parts.append(f"      <FullName>{xml_escape(name)}</FullName>")
    parts.append("    </ItemQueryRq>")
    return wrap_request("\n".join(parts))


def _build_terms_query(names: Sequence[str]) -> str:
    parts: list[str] = ['    <TermsQueryRq requestID="1">']
    for name in names:
        parts.append(f"      <Name>{xml_escape(name)}</Name>")
    parts.append("    </TermsQueryRq>")
    return wrap_request("\n".join(parts))


def _collect_names_from_response(response_xml: str) -> set[str]:
    """Walk the parsed response and return every Name/FullName text we find."""
    root = etree.fromstring(response_xml.encode("utf-8"))
    names: set[str] = set()
    for tag in ("FullName", "Name"):
        for elem in root.iter(tag):
            if elem.text is not None:
                stripped = elem.text.strip()
                if stripped:
                    names.add(stripped)
    return names


def _verify_bucket(
    ctx: Context,
    names: Sequence[str],
    request_builder: Callable[[Sequence[str]], str],
) -> tuple[set[str], set[str]]:
    if not names:
        return set(), set()
    requested: set[str] = set(names)
    request_xml = request_builder(list(requested))
    try:
        with ctx.connection_factory() as conn:
            response_xml = conn.send(request_xml)
        check_response_status(response_xml)
    except QBEntityNotFound:
        return set(), set(requested)
    found_names = _collect_names_from_response(response_xml) & requested
    missing = requested - found_names
    return found_names, missing


def verify_entities(
    ctx: Context,
    *,
    customers: Sequence[str] = (),
    vendors: Sequence[str] = (),
    items: Sequence[str] = (),
    terms: Sequence[str] = (),
) -> VerifyResult:
    result = VerifyResult()
    for bucket, names, builder in (
        ("customer", customers, _build_customer_query),
        ("vendor", vendors, _build_vendor_query),
        ("item", items, _build_item_query),
        ("terms", terms, _build_terms_query),
    ):
        found, missing = _verify_bucket(ctx, names, builder)
        result.found[bucket] = found
        result.missing[bucket] = missing
    return result
