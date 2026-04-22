"""
QBXML builder and parser for QuickBooks Desktop Purchase Orders.

Provides functions to build PurchaseOrderQueryRq / PurchaseOrderAddRq /
PurchaseOrderModRq request bodies, and to parse the corresponding responses
into :class:`qb_cli.models.purchase_order.PurchaseOrder` instances.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import Optional

from lxml import etree

from qb_cli.models.purchase_order import PurchaseOrder, PurchaseOrderLineItem
from qb_cli.models.shared import Address, Ref
from qb_cli.qbxml.common import format_qb_date, parse_address, xml_escape
from qb_cli.qbxml.envelope import wrap_request


# ---------------------------------------------------------------------------
# Builders
# ---------------------------------------------------------------------------

_ADDRESS_TAGS: tuple[str, ...] = (
    "Addr1", "Addr2", "Addr3", "Addr4", "Addr5",
    "City", "State", "PostalCode", "Country", "Note",
)


def build_query(
    *,
    ref_numbers: Optional[list[str]] = None,
    txn_ids: Optional[list[str]] = None,
    date_from: Optional[date] = None,
    date_to: Optional[date] = None,
    include_line_items: bool = True,
    max_returned: Optional[int] = None,
    iterator_id: Optional[str] = None,
) -> str:
    """Build a PurchaseOrderQueryRq QBXML request string.

    Ordering rules enforced:
      1. TxnID
      2. RefNumber
      3. MaxReturned (only if no TxnID/RefNumber filter is present)
      4. TxnDateRangeFilter
      5. IncludeLineItems
    """
    iterator_attr = ""
    if iterator_id is not None:
        iterator_attr = (
            f' iterator="Continue" iteratorID="{xml_escape(iterator_id)}"'
        )

    parts: list[str] = []
    parts.append(f'    <PurchaseOrderQueryRq requestID="1"{iterator_attr}>')

    if txn_ids:
        for tid in txn_ids:
            parts.append(f'      <TxnID>{xml_escape(tid)}</TxnID>')

    if ref_numbers:
        for ref in ref_numbers:
            parts.append(f'      <RefNumber>{xml_escape(ref)}</RefNumber>')

    # MaxReturned is not allowed with RefNumber/TxnID filters.
    if max_returned is not None and not txn_ids and not ref_numbers:
        parts.append(f'      <MaxReturned>{int(max_returned)}</MaxReturned>')

    if date_from is not None or date_to is not None:
        parts.append('      <TxnDateRangeFilter>')
        if date_from is not None:
            parts.append(
                f'        <FromTxnDate>{format_qb_date(date_from)}</FromTxnDate>'
            )
        if date_to is not None:
            parts.append(
                f'        <ToTxnDate>{format_qb_date(date_to)}</ToTxnDate>'
            )
        parts.append('      </TxnDateRangeFilter>')

    include_flag = "true" if include_line_items else "false"
    parts.append(f'      <IncludeLineItems>{include_flag}</IncludeLineItems>')

    parts.append('    </PurchaseOrderQueryRq>')

    return wrap_request("\n".join(parts))


def _render_ref(ref: Ref, tag: str, indent: str) -> list[str]:
    lines: list[str] = [f'{indent}<{tag}>']
    if ref.list_id:
        lines.append(f'{indent}  <ListID>{xml_escape(ref.list_id)}</ListID>')
    elif ref.full_name:
        lines.append(f'{indent}  <FullName>{xml_escape(ref.full_name)}</FullName>')
    lines.append(f'{indent}</{tag}>')
    return lines


def _render_address(addr: Address, tag: str, indent: str) -> list[str]:
    lines: list[str] = [f'{indent}<{tag}>']
    payload = addr.model_dump(by_alias=True, exclude_none=True)
    for qb_tag in _ADDRESS_TAGS:
        value = payload.get(qb_tag)
        if value:
            lines.append(
                f'{indent}  <{qb_tag}>{xml_escape(str(value))}</{qb_tag}>'
            )
    lines.append(f'{indent}</{tag}>')
    return lines


def _render_decimal(value: Decimal) -> str:
    return str(value)


def _render_bool(value: bool) -> str:
    return "true" if value else "false"


def _render_line_item(
    li: PurchaseOrderLineItem, *, mode: str, indent: str
) -> list[str]:
    """Render a line item.

    mode is either "add" or "mod". In mod mode, TxnLineID is emitted first
    (using "-1" for new lines if absent).
    """
    if mode == "add":
        wrapper = "PurchaseOrderLineAdd"
    elif mode == "mod":
        wrapper = "PurchaseOrderLineMod"
    else:
        raise ValueError(f"Invalid line-item mode: {mode!r}")

    lines: list[str] = [f'{indent}<{wrapper}>']
    child_indent = indent + "  "

    if mode == "mod":
        txn_line_id = li.txn_line_id if li.txn_line_id else "-1"
        lines.append(
            f'{child_indent}<TxnLineID>{xml_escape(txn_line_id)}</TxnLineID>'
        )

    if li.item_ref is not None:
        lines.extend(_render_ref(li.item_ref, "ItemRef", child_indent))

    if li.manufacturer_part_number:
        lines.append(
            f'{child_indent}<ManufacturerPartNumber>'
            f'{xml_escape(li.manufacturer_part_number)}'
            f'</ManufacturerPartNumber>'
        )

    if li.description:
        lines.append(
            f'{child_indent}<Desc>{xml_escape(li.description)}</Desc>'
        )

    if li.quantity is not None:
        lines.append(
            f'{child_indent}<Quantity>{_render_decimal(li.quantity)}</Quantity>'
        )

    if li.unit_of_measure:
        lines.append(
            f'{child_indent}<UnitOfMeasure>'
            f'{xml_escape(li.unit_of_measure)}'
            f'</UnitOfMeasure>'
        )

    if li.rate is not None:
        lines.append(
            f'{child_indent}<Rate>{_render_decimal(li.rate)}</Rate>'
        )

    if li.class_ref is not None:
        lines.extend(_render_ref(li.class_ref, "ClassRef", child_indent))

    if li.amount is not None:
        lines.append(
            f'{child_indent}<Amount>{_render_decimal(li.amount)}</Amount>'
        )

    if li.inventory_site_ref is not None:
        lines.extend(
            _render_ref(li.inventory_site_ref, "InventorySiteRef", child_indent)
        )

    if li.inventory_site_location_ref is not None:
        lines.extend(
            _render_ref(
                li.inventory_site_location_ref,
                "InventorySiteLocationRef",
                child_indent,
            )
        )

    if li.customer_ref is not None:
        lines.extend(_render_ref(li.customer_ref, "CustomerRef", child_indent))

    if li.service_date is not None:
        lines.append(
            f'{child_indent}<ServiceDate>'
            f'{format_qb_date(li.service_date)}'
            f'</ServiceDate>'
        )

    if li.sales_tax_code_ref is not None:
        lines.extend(
            _render_ref(li.sales_tax_code_ref, "SalesTaxCodeRef", child_indent)
        )

    if li.is_manually_closed is not None:
        lines.append(
            f'{child_indent}<IsManuallyClosed>'
            f'{_render_bool(li.is_manually_closed)}'
            f'</IsManuallyClosed>'
        )

    if li.other1:
        lines.append(
            f'{child_indent}<Other1>{xml_escape(li.other1)}</Other1>'
        )

    if li.other2:
        lines.append(
            f'{child_indent}<Other2>{xml_escape(li.other2)}</Other2>'
        )

    lines.append(f'{indent}</{wrapper}>')
    return lines


def _render_header_fields(
    entity: PurchaseOrder, *, indent: str, mode: str
) -> list[str]:
    """Emit the header-level (non-line-item) fields in canonical order.

    Canonical QBXML order (Add/Mod):
      VendorRef, ClassRef, TemplateRef (add-only)?, TxnDate, RefNumber,
      VendorAddress, ShipAddress, TermsRef, DueDate, ExpectedDate,
      ShipMethodRef, FOB, Memo, VendorMsg, IsToBePrinted, IsToBeEmailed,
      IsTaxIncluded, SalesTaxCodeRef, Other1, Other2, ExchangeRate,
      IsManuallyClosed
    """
    lines: list[str] = []

    if entity.vendor_ref is not None:
        lines.extend(_render_ref(entity.vendor_ref, "VendorRef", indent))

    if entity.class_ref is not None:
        lines.extend(_render_ref(entity.class_ref, "ClassRef", indent))

    if entity.template_ref is not None:
        lines.extend(_render_ref(entity.template_ref, "TemplateRef", indent))

    if entity.txn_date is not None:
        lines.append(
            f'{indent}<TxnDate>{format_qb_date(entity.txn_date)}</TxnDate>'
        )

    if entity.ref_number:
        lines.append(
            f'{indent}<RefNumber>{xml_escape(entity.ref_number)}</RefNumber>'
        )

    if entity.vendor_address is not None:
        lines.extend(
            _render_address(entity.vendor_address, "VendorAddress", indent)
        )

    if entity.ship_address is not None:
        lines.extend(
            _render_address(entity.ship_address, "ShipAddress", indent)
        )

    if entity.terms_ref is not None:
        lines.extend(_render_ref(entity.terms_ref, "TermsRef", indent))

    if entity.due_date is not None:
        lines.append(
            f'{indent}<DueDate>{format_qb_date(entity.due_date)}</DueDate>'
        )

    if entity.expected_date is not None:
        lines.append(
            f'{indent}<ExpectedDate>'
            f'{format_qb_date(entity.expected_date)}'
            f'</ExpectedDate>'
        )

    if entity.ship_method_ref is not None:
        lines.extend(
            _render_ref(entity.ship_method_ref, "ShipMethodRef", indent)
        )

    if entity.fob:
        lines.append(f'{indent}<FOB>{xml_escape(entity.fob)}</FOB>')

    if entity.memo:
        lines.append(f'{indent}<Memo>{xml_escape(entity.memo)}</Memo>')

    if entity.vendor_msg:
        lines.append(
            f'{indent}<VendorMsg>{xml_escape(entity.vendor_msg)}</VendorMsg>'
        )

    if entity.is_to_be_printed is not None:
        lines.append(
            f'{indent}<IsToBePrinted>'
            f'{_render_bool(entity.is_to_be_printed)}'
            f'</IsToBePrinted>'
        )

    if entity.is_to_be_emailed is not None:
        lines.append(
            f'{indent}<IsToBeEmailed>'
            f'{_render_bool(entity.is_to_be_emailed)}'
            f'</IsToBeEmailed>'
        )

    if entity.is_tax_included is not None:
        lines.append(
            f'{indent}<IsTaxIncluded>'
            f'{_render_bool(entity.is_tax_included)}'
            f'</IsTaxIncluded>'
        )

    if entity.sales_tax_code_ref is not None:
        lines.extend(
            _render_ref(entity.sales_tax_code_ref, "SalesTaxCodeRef", indent)
        )

    if entity.other1:
        lines.append(f'{indent}<Other1>{xml_escape(entity.other1)}</Other1>')

    if entity.other2:
        lines.append(f'{indent}<Other2>{xml_escape(entity.other2)}</Other2>')

    if entity.exchange_rate is not None:
        lines.append(
            f'{indent}<ExchangeRate>'
            f'{_render_decimal(entity.exchange_rate)}'
            f'</ExchangeRate>'
        )

    if entity.is_manually_closed is not None:
        lines.append(
            f'{indent}<IsManuallyClosed>'
            f'{_render_bool(entity.is_manually_closed)}'
            f'</IsManuallyClosed>'
        )

    # Silence "unused" warning when mode affects ordering only conceptually.
    _ = mode
    return lines


def build_add(entity: PurchaseOrder) -> str:
    """Build a PurchaseOrderAddRq QBXML request string."""
    inner_indent = "        "
    parts: list[str] = []
    parts.append('    <PurchaseOrderAddRq>')
    parts.append('      <PurchaseOrderAdd>')

    parts.extend(
        _render_header_fields(entity, indent=inner_indent, mode="add")
    )

    for li in entity.line_items:
        parts.extend(_render_line_item(li, mode="add", indent=inner_indent))

    parts.append('      </PurchaseOrderAdd>')
    parts.append('    </PurchaseOrderAddRq>')

    return wrap_request("\n".join(parts))


def build_mod(entity: PurchaseOrder) -> str:
    """Build a PurchaseOrderModRq QBXML request string.

    Requires both ``txn_id`` and ``edit_sequence``.
    """
    if not entity.edit_sequence:
        raise ValueError("build_mod requires EditSequence")
    if not entity.txn_id:
        raise ValueError("build_mod requires TxnID")

    inner_indent = "        "
    parts: list[str] = []
    parts.append('    <PurchaseOrderModRq>')
    parts.append('      <PurchaseOrderMod>')

    parts.append(f'{inner_indent}<TxnID>{xml_escape(entity.txn_id)}</TxnID>')
    parts.append(
        f'{inner_indent}<EditSequence>'
        f'{xml_escape(entity.edit_sequence)}'
        f'</EditSequence>'
    )

    parts.extend(
        _render_header_fields(entity, indent=inner_indent, mode="mod")
    )

    for li in entity.line_items:
        parts.extend(_render_line_item(li, mode="mod", indent=inner_indent))

    parts.append('      </PurchaseOrderMod>')
    parts.append('    </PurchaseOrderModRq>')

    return wrap_request("\n".join(parts))


# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------


def _parse_ref(elem: Optional[etree._Element]) -> Optional[dict[str, str]]:
    if elem is None:
        return None
    result: dict[str, str] = {}
    list_id_el = elem.find("ListID")
    if list_id_el is not None and list_id_el.text:
        result["ListID"] = list_id_el.text.strip()
    full_name_el = elem.find("FullName")
    if full_name_el is not None and full_name_el.text:
        result["FullName"] = full_name_el.text.strip()
    if not result:
        return None
    return result


def _text(elem: Optional[etree._Element], tag: str) -> Optional[str]:
    if elem is None:
        return None
    child = elem.find(tag)
    if child is None or child.text is None:
        return None
    stripped = child.text.strip()
    return stripped or None


def _text_decimal(elem: Optional[etree._Element], tag: str) -> Optional[Decimal]:
    raw = _text(elem, tag)
    if raw is None:
        return None
    return Decimal(raw)


def _text_bool(elem: Optional[etree._Element], tag: str) -> Optional[bool]:
    raw = _text(elem, tag)
    if raw is None:
        return None
    return raw.lower() == "true"


def _text_int(elem: Optional[etree._Element], tag: str) -> Optional[int]:
    raw = _text(elem, tag)
    if raw is None:
        return None
    return int(raw)


def _text_date(elem: Optional[etree._Element], tag: str) -> Optional[date]:
    raw = _text(elem, tag)
    if raw is None:
        return None
    return date.fromisoformat(raw)


def _parse_address_payload(
    elem: Optional[etree._Element],
) -> Optional[dict[str, str]]:
    addr = parse_address(elem)
    if addr is None:
        return None
    return addr.model_dump(by_alias=True, exclude_none=True)


def _parse_line_item(line_elem: etree._Element) -> PurchaseOrderLineItem:
    payload: dict[str, object] = {}

    txn_line_id = _text(line_elem, "TxnLineID")
    if txn_line_id is not None:
        payload["TxnLineID"] = txn_line_id

    item_ref = _parse_ref(line_elem.find("ItemRef"))
    if item_ref is not None:
        payload["ItemRef"] = item_ref

    mpn = _text(line_elem, "ManufacturerPartNumber")
    if mpn is not None:
        payload["ManufacturerPartNumber"] = mpn

    desc = _text(line_elem, "Desc")
    if desc is not None:
        payload["Desc"] = desc

    quantity = _text_decimal(line_elem, "Quantity")
    if quantity is not None:
        payload["Quantity"] = quantity

    unit = _text(line_elem, "UnitOfMeasure")
    if unit is not None:
        payload["UnitOfMeasure"] = unit

    rate = _text_decimal(line_elem, "Rate")
    if rate is not None:
        payload["Rate"] = rate

    class_ref = _parse_ref(line_elem.find("ClassRef"))
    if class_ref is not None:
        payload["ClassRef"] = class_ref

    amount = _text_decimal(line_elem, "Amount")
    if amount is not None:
        payload["Amount"] = amount

    inventory_site = _parse_ref(line_elem.find("InventorySiteRef"))
    if inventory_site is not None:
        payload["InventorySiteRef"] = inventory_site

    inventory_site_loc = _parse_ref(line_elem.find("InventorySiteLocationRef"))
    if inventory_site_loc is not None:
        payload["InventorySiteLocationRef"] = inventory_site_loc

    customer_ref = _parse_ref(line_elem.find("CustomerRef"))
    if customer_ref is not None:
        payload["CustomerRef"] = customer_ref

    service_date = _text_date(line_elem, "ServiceDate")
    if service_date is not None:
        payload["ServiceDate"] = service_date

    sales_tax = _parse_ref(line_elem.find("SalesTaxCodeRef"))
    if sales_tax is not None:
        payload["SalesTaxCodeRef"] = sales_tax

    received = _text_decimal(line_elem, "ReceivedQuantity")
    if received is not None:
        payload["ReceivedQuantity"] = received

    is_closed = _text_bool(line_elem, "IsManuallyClosed")
    if is_closed is not None:
        payload["IsManuallyClosed"] = is_closed

    other1 = _text(line_elem, "Other1")
    if other1 is not None:
        payload["Other1"] = other1

    other2 = _text(line_elem, "Other2")
    if other2 is not None:
        payload["Other2"] = other2

    return PurchaseOrderLineItem.model_validate(payload)


def _parse_purchase_order_ret(ret_elem: etree._Element) -> PurchaseOrder:
    payload: dict[str, object] = {}

    for tag in ("TxnID", "TimeCreated", "TimeModified", "EditSequence"):
        value = _text(ret_elem, tag)
        if value is not None:
            payload[tag] = value

    txn_number = _text_int(ret_elem, "TxnNumber")
    if txn_number is not None:
        payload["TxnNumber"] = txn_number

    vendor_ref = _parse_ref(ret_elem.find("VendorRef"))
    if vendor_ref is not None:
        payload["VendorRef"] = vendor_ref

    class_ref = _parse_ref(ret_elem.find("ClassRef"))
    if class_ref is not None:
        payload["ClassRef"] = class_ref

    template_ref = _parse_ref(ret_elem.find("TemplateRef"))
    if template_ref is not None:
        payload["TemplateRef"] = template_ref

    txn_date = _text_date(ret_elem, "TxnDate")
    if txn_date is not None:
        payload["TxnDate"] = txn_date

    ref_number = _text(ret_elem, "RefNumber")
    if ref_number is not None:
        payload["RefNumber"] = ref_number

    vendor_addr = _parse_address_payload(ret_elem.find("VendorAddress"))
    if vendor_addr is not None:
        payload["VendorAddress"] = vendor_addr

    ship_addr = _parse_address_payload(ret_elem.find("ShipAddress"))
    if ship_addr is not None:
        payload["ShipAddress"] = ship_addr

    terms_ref = _parse_ref(ret_elem.find("TermsRef"))
    if terms_ref is not None:
        payload["TermsRef"] = terms_ref

    due_date = _text_date(ret_elem, "DueDate")
    if due_date is not None:
        payload["DueDate"] = due_date

    expected_date = _text_date(ret_elem, "ExpectedDate")
    if expected_date is not None:
        payload["ExpectedDate"] = expected_date

    ship_method_ref = _parse_ref(ret_elem.find("ShipMethodRef"))
    if ship_method_ref is not None:
        payload["ShipMethodRef"] = ship_method_ref

    fob = _text(ret_elem, "FOB")
    if fob is not None:
        payload["FOB"] = fob

    memo = _text(ret_elem, "Memo")
    if memo is not None:
        payload["Memo"] = memo

    vendor_msg = _text(ret_elem, "VendorMsg")
    if vendor_msg is not None:
        payload["VendorMsg"] = vendor_msg

    for tag in ("IsToBePrinted", "IsToBeEmailed", "IsTaxIncluded",
                "IsManuallyClosed", "IsFullyReceived"):
        bvalue = _text_bool(ret_elem, tag)
        if bvalue is not None:
            payload[tag] = bvalue

    sales_tax = _parse_ref(ret_elem.find("SalesTaxCodeRef"))
    if sales_tax is not None:
        payload["SalesTaxCodeRef"] = sales_tax

    other1 = _text(ret_elem, "Other1")
    if other1 is not None:
        payload["Other1"] = other1

    other2 = _text(ret_elem, "Other2")
    if other2 is not None:
        payload["Other2"] = other2

    exchange_rate = _text_decimal(ret_elem, "ExchangeRate")
    if exchange_rate is not None:
        payload["ExchangeRate"] = exchange_rate

    for tag in ("Subtotal", "SalesTaxTotal", "TotalAmount"):
        dvalue = _text_decimal(ret_elem, tag)
        if dvalue is not None:
            payload[tag] = dvalue

    line_items: list[PurchaseOrderLineItem] = [
        _parse_line_item(line_elem)
        for line_elem in ret_elem.findall("PurchaseOrderLineRet")
    ]
    payload["PurchaseOrderLineRet"] = line_items

    return PurchaseOrder.model_validate(payload)


def parse_query_response(xml: str) -> list[PurchaseOrder]:
    """Parse a PurchaseOrderQueryRs response into a list of PurchaseOrder."""
    root = etree.fromstring(xml.encode("utf-8"))
    return [
        _parse_purchase_order_ret(ret)
        for ret in root.iter("PurchaseOrderRet")
    ]


def parse_add_response(xml: str) -> PurchaseOrder:
    """Parse a PurchaseOrderAddRs response into a single PurchaseOrder."""
    root = etree.fromstring(xml.encode("utf-8"))
    ret = root.find(".//PurchaseOrderAddRs/PurchaseOrderRet")
    if ret is None:
        raise ValueError("PurchaseOrderAddRs/PurchaseOrderRet not found")
    return _parse_purchase_order_ret(ret)


def parse_mod_response(xml: str) -> PurchaseOrder:
    """Parse a PurchaseOrderModRs response into a single PurchaseOrder."""
    root = etree.fromstring(xml.encode("utf-8"))
    ret = root.find(".//PurchaseOrderModRs/PurchaseOrderRet")
    if ret is None:
        raise ValueError("PurchaseOrderModRs/PurchaseOrderRet not found")
    return _parse_purchase_order_ret(ret)
