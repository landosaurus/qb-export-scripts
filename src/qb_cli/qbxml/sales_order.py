from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import Optional

from lxml import etree

from qb_cli.models.shared import Address, Ref
from qb_cli.models.sales_order import SalesOrder, SalesOrderLineItem
from qb_cli.qbxml.common import format_qb_date, parse_address, xml_escape
from qb_cli.qbxml.envelope import wrap_request


# -----------------------------------------------------------------------------
# Builders
# -----------------------------------------------------------------------------


def _render_ref(tag: str, ref: Optional[Ref], indent: str) -> list[str]:
    if ref is None:
        return []
    lines: list[str] = [f"{indent}<{tag}>"]
    if ref.list_id:
        lines.append(f"{indent}  <ListID>{xml_escape(ref.list_id)}</ListID>")
    if ref.full_name:
        lines.append(f"{indent}  <FullName>{xml_escape(ref.full_name)}</FullName>")
    lines.append(f"{indent}</{tag}>")
    return lines


def _render_address(tag: str, addr: Optional[Address], indent: str) -> list[str]:
    if addr is None:
        return []
    lines: list[str] = [f"{indent}<{tag}>"]
    for field_name, alias in (
        ("addr1", "Addr1"),
        ("addr2", "Addr2"),
        ("addr3", "Addr3"),
        ("addr4", "Addr4"),
        ("addr5", "Addr5"),
        ("city", "City"),
        ("state", "State"),
        ("postal_code", "PostalCode"),
        ("country", "Country"),
        ("note", "Note"),
    ):
        value = getattr(addr, field_name)
        if value:
            lines.append(f"{indent}  <{alias}>{xml_escape(value)}</{alias}>")
    lines.append(f"{indent}</{tag}>")
    return lines


def _bool_str(value: bool) -> str:
    return "true" if value else "false"


def _decimal_str(value: Decimal) -> str:
    return str(value)


def build_query(
    *,
    ref_numbers: list[str] | None = None,
    txn_ids: list[str] | None = None,
    date_from: date | None = None,
    date_to: date | None = None,
    include_line_items: bool = True,
    max_returned: int | None = None,
    iterator_id: str | None = None,
) -> str:
    """Build a SalesOrderQueryRq QBXML request."""
    # Build attributes for the Rq element. Iterator handling mirrors invoice.py.
    attrs = 'requestID="1"'
    if iterator_id is not None:
        attrs += f' iterator="Continue" iteratorID="{xml_escape(iterator_id)}"'
    elif max_returned is not None and not ref_numbers and not txn_ids:
        # Implicit iterator=Start when only max_returned is provided
        attrs += ' iterator="Start"'

    lines: list[str] = [f"    <SalesOrderQueryRq {attrs}>"]

    # Specific lookups: TxnID, RefNumber (before MaxReturned & filters)
    if txn_ids:
        for tid in txn_ids:
            lines.append(f"      <TxnID>{xml_escape(tid)}</TxnID>")

    if ref_numbers:
        for ref in ref_numbers:
            lines.append(f"      <RefNumber>{xml_escape(ref)}</RefNumber>")

    # MaxReturned (not allowed with RefNumber/TxnID filters)
    if max_returned is not None and not ref_numbers and not txn_ids:
        lines.append(f"      <MaxReturned>{max_returned}</MaxReturned>")

    # TxnDateRangeFilter comes BEFORE IncludeLineItems
    if date_from is not None or date_to is not None:
        lines.append("      <TxnDateRangeFilter>")
        if date_from is not None:
            lines.append(
                f"        <FromTxnDate>{format_qb_date(date_from)}</FromTxnDate>"
            )
        if date_to is not None:
            lines.append(
                f"        <ToTxnDate>{format_qb_date(date_to)}</ToTxnDate>"
            )
        lines.append("      </TxnDateRangeFilter>")

    # IncludeLineItems
    lines.append(
        f"      <IncludeLineItems>{_bool_str(include_line_items)}</IncludeLineItems>"
    )

    lines.append("    </SalesOrderQueryRq>")
    return wrap_request("\n".join(lines))


def _render_line_item_add(item: SalesOrderLineItem, indent: str = "        ") -> list[str]:
    lines: list[str] = [f"{indent}<SalesOrderLineAdd>"]
    inner = indent + "  "

    if item.item_ref is not None:
        lines.extend(_render_ref("ItemRef", item.item_ref, inner))
    if item.description is not None:
        lines.append(f"{inner}<Desc>{xml_escape(item.description)}</Desc>")
    if item.quantity is not None:
        lines.append(f"{inner}<Quantity>{_decimal_str(item.quantity)}</Quantity>")
    if item.unit_of_measure is not None:
        lines.append(
            f"{inner}<UnitOfMeasure>{xml_escape(item.unit_of_measure)}</UnitOfMeasure>"
        )
    if item.rate is not None:
        lines.append(f"{inner}<Rate>{_decimal_str(item.rate)}</Rate>")
    if item.rate_percent is not None:
        lines.append(
            f"{inner}<RatePercent>{_decimal_str(item.rate_percent)}</RatePercent>"
        )
    if item.price_level_ref is not None:
        lines.extend(_render_ref("PriceLevelRef", item.price_level_ref, inner))
    if item.class_ref is not None:
        lines.extend(_render_ref("ClassRef", item.class_ref, inner))
    if item.amount is not None:
        lines.append(f"{inner}<Amount>{_decimal_str(item.amount)}</Amount>")
    if item.inventory_site_ref is not None:
        lines.extend(_render_ref("InventorySiteRef", item.inventory_site_ref, inner))
    if item.inventory_site_location_ref is not None:
        lines.extend(
            _render_ref(
                "InventorySiteLocationRef",
                item.inventory_site_location_ref,
                inner,
            )
        )
    if item.serial_number is not None:
        lines.append(
            f"{inner}<SerialNumber>{xml_escape(item.serial_number)}</SerialNumber>"
        )
    if item.lot_number is not None:
        lines.append(f"{inner}<LotNumber>{xml_escape(item.lot_number)}</LotNumber>")
    if item.sales_tax_code_ref is not None:
        lines.extend(_render_ref("SalesTaxCodeRef", item.sales_tax_code_ref, inner))
    if item.is_taxable is not None:
        lines.append(f"{inner}<IsTaxable>{_bool_str(item.is_taxable)}</IsTaxable>")
    if item.is_manually_closed is not None:
        lines.append(
            f"{inner}<IsManuallyClosed>{_bool_str(item.is_manually_closed)}</IsManuallyClosed>"
        )
    if item.other1 is not None:
        lines.append(f"{inner}<Other1>{xml_escape(item.other1)}</Other1>")
    if item.other2 is not None:
        lines.append(f"{inner}<Other2>{xml_escape(item.other2)}</Other2>")

    lines.append(f"{indent}</SalesOrderLineAdd>")
    return lines


def _render_line_item_mod(item: SalesOrderLineItem, indent: str = "        ") -> list[str]:
    # If no txn_line_id, treat as new line add (still under SalesOrderLineMod with TxnLineID -1)
    tag = "SalesOrderLineMod"
    lines: list[str] = [f"{indent}<{tag}>"]
    inner = indent + "  "

    txn_line_id = item.txn_line_id if item.txn_line_id else "-1"
    lines.append(f"{inner}<TxnLineID>{xml_escape(txn_line_id)}</TxnLineID>")

    if item.item_ref is not None:
        lines.extend(_render_ref("ItemRef", item.item_ref, inner))
    if item.description is not None:
        lines.append(f"{inner}<Desc>{xml_escape(item.description)}</Desc>")
    if item.quantity is not None:
        lines.append(f"{inner}<Quantity>{_decimal_str(item.quantity)}</Quantity>")
    if item.unit_of_measure is not None:
        lines.append(
            f"{inner}<UnitOfMeasure>{xml_escape(item.unit_of_measure)}</UnitOfMeasure>"
        )
    if item.rate is not None:
        lines.append(f"{inner}<Rate>{_decimal_str(item.rate)}</Rate>")
    if item.rate_percent is not None:
        lines.append(
            f"{inner}<RatePercent>{_decimal_str(item.rate_percent)}</RatePercent>"
        )
    if item.price_level_ref is not None:
        lines.extend(_render_ref("PriceLevelRef", item.price_level_ref, inner))
    if item.class_ref is not None:
        lines.extend(_render_ref("ClassRef", item.class_ref, inner))
    if item.amount is not None:
        lines.append(f"{inner}<Amount>{_decimal_str(item.amount)}</Amount>")
    if item.inventory_site_ref is not None:
        lines.extend(_render_ref("InventorySiteRef", item.inventory_site_ref, inner))
    if item.inventory_site_location_ref is not None:
        lines.extend(
            _render_ref(
                "InventorySiteLocationRef",
                item.inventory_site_location_ref,
                inner,
            )
        )
    if item.serial_number is not None:
        lines.append(
            f"{inner}<SerialNumber>{xml_escape(item.serial_number)}</SerialNumber>"
        )
    if item.lot_number is not None:
        lines.append(f"{inner}<LotNumber>{xml_escape(item.lot_number)}</LotNumber>")
    if item.sales_tax_code_ref is not None:
        lines.extend(_render_ref("SalesTaxCodeRef", item.sales_tax_code_ref, inner))
    if item.is_taxable is not None:
        lines.append(f"{inner}<IsTaxable>{_bool_str(item.is_taxable)}</IsTaxable>")
    if item.is_manually_closed is not None:
        lines.append(
            f"{inner}<IsManuallyClosed>{_bool_str(item.is_manually_closed)}</IsManuallyClosed>"
        )
    if item.other1 is not None:
        lines.append(f"{inner}<Other1>{xml_escape(item.other1)}</Other1>")
    if item.other2 is not None:
        lines.append(f"{inner}<Other2>{xml_escape(item.other2)}</Other2>")

    lines.append(f"{indent}</{tag}>")
    return lines


def _render_header_common(entity: SalesOrder, inner: str) -> list[str]:
    """Render header fields shared between Add and Mod (excluding identity)."""
    lines: list[str] = []

    if entity.customer_ref is not None:
        lines.extend(_render_ref("CustomerRef", entity.customer_ref, inner))
    if entity.class_ref is not None:
        lines.extend(_render_ref("ClassRef", entity.class_ref, inner))
    if entity.template_ref is not None:
        lines.extend(_render_ref("TemplateRef", entity.template_ref, inner))
    if entity.txn_date is not None:
        lines.append(f"{inner}<TxnDate>{xml_escape(entity.txn_date)}</TxnDate>")
    if entity.ref_number is not None:
        lines.append(f"{inner}<RefNumber>{xml_escape(entity.ref_number)}</RefNumber>")
    if entity.bill_address is not None:
        lines.extend(_render_address("BillAddress", entity.bill_address, inner))
    if entity.ship_address is not None:
        lines.extend(_render_address("ShipAddress", entity.ship_address, inner))
    if entity.po_number is not None:
        lines.append(f"{inner}<PONumber>{xml_escape(entity.po_number)}</PONumber>")
    if entity.terms_ref is not None:
        lines.extend(_render_ref("TermsRef", entity.terms_ref, inner))
    if entity.due_date is not None:
        lines.append(f"{inner}<DueDate>{xml_escape(entity.due_date)}</DueDate>")
    if entity.sales_rep_ref is not None:
        lines.extend(_render_ref("SalesRepRef", entity.sales_rep_ref, inner))
    if entity.fob is not None:
        lines.append(f"{inner}<FOB>{xml_escape(entity.fob)}</FOB>")
    if entity.ship_date is not None:
        lines.append(f"{inner}<ShipDate>{xml_escape(entity.ship_date)}</ShipDate>")
    if entity.ship_method_ref is not None:
        lines.extend(_render_ref("ShipMethodRef", entity.ship_method_ref, inner))
    if entity.item_sales_tax_ref is not None:
        lines.extend(_render_ref("ItemSalesTaxRef", entity.item_sales_tax_ref, inner))
    if entity.memo is not None:
        lines.append(f"{inner}<Memo>{xml_escape(entity.memo)}</Memo>")
    if entity.is_manually_closed is not None:
        lines.append(
            f"{inner}<IsManuallyClosed>{_bool_str(entity.is_manually_closed)}</IsManuallyClosed>"
        )
    if entity.is_to_be_printed is not None:
        lines.append(
            f"{inner}<IsToBePrinted>{_bool_str(entity.is_to_be_printed)}</IsToBePrinted>"
        )
    if entity.is_to_be_emailed is not None:
        lines.append(
            f"{inner}<IsToBeEmailed>{_bool_str(entity.is_to_be_emailed)}</IsToBeEmailed>"
        )
    if entity.is_tax_included is not None:
        lines.append(
            f"{inner}<IsTaxIncluded>{_bool_str(entity.is_tax_included)}</IsTaxIncluded>"
        )
    if entity.customer_sales_tax_code_ref is not None:
        lines.extend(
            _render_ref(
                "CustomerSalesTaxCodeRef", entity.customer_sales_tax_code_ref, inner
            )
        )
    if entity.other is not None:
        lines.append(f"{inner}<Other>{xml_escape(entity.other)}</Other>")
    if entity.exchange_rate is not None:
        lines.append(
            f"{inner}<ExchangeRate>{_decimal_str(entity.exchange_rate)}</ExchangeRate>"
        )

    return lines


def build_add(entity: SalesOrder) -> str:
    """Build a SalesOrderAddRq QBXML request from a SalesOrder model."""
    lines: list[str] = [
        '    <SalesOrderAddRq requestID="1">',
        "      <SalesOrderAdd>",
    ]
    inner = "        "
    lines.extend(_render_header_common(entity, inner))

    for item in entity.line_items:
        lines.extend(_render_line_item_add(item, inner))

    lines.append("      </SalesOrderAdd>")
    lines.append("    </SalesOrderAddRq>")
    return wrap_request("\n".join(lines))


def build_mod(entity: SalesOrder) -> str:
    """Build a SalesOrderModRq QBXML request from a SalesOrder model."""
    if not entity.edit_sequence:
        raise ValueError("build_mod requires EditSequence")
    if not entity.txn_id:
        raise ValueError("build_mod requires TxnID")

    lines: list[str] = [
        '    <SalesOrderModRq requestID="1">',
        "      <SalesOrderMod>",
    ]
    inner = "        "
    lines.append(f"{inner}<TxnID>{xml_escape(entity.txn_id)}</TxnID>")
    lines.append(
        f"{inner}<EditSequence>{xml_escape(entity.edit_sequence)}</EditSequence>"
    )

    lines.extend(_render_header_common(entity, inner))

    for item in entity.line_items:
        lines.extend(_render_line_item_mod(item, inner))

    lines.append("      </SalesOrderMod>")
    lines.append("    </SalesOrderModRq>")
    return wrap_request("\n".join(lines))


# -----------------------------------------------------------------------------
# Parsers
# -----------------------------------------------------------------------------


def _parse_ref(elem: Optional[etree._Element]) -> Optional[dict[str, str]]:
    if elem is None:
        return None
    payload: dict[str, str] = {}
    list_id = elem.findtext("ListID")
    full_name = elem.findtext("FullName")
    if list_id:
        payload["ListID"] = list_id.strip()
    if full_name:
        payload["FullName"] = full_name.strip()
    if not payload:
        return None
    return payload


def _parse_line_item(elem: etree._Element) -> SalesOrderLineItem:
    payload: dict[str, object] = {}

    txn_line_id = elem.findtext("TxnLineID")
    if txn_line_id:
        payload["TxnLineID"] = txn_line_id.strip()

    item_ref = _parse_ref(elem.find("ItemRef"))
    if item_ref is not None:
        payload["ItemRef"] = item_ref

    desc = elem.findtext("Desc")
    if desc:
        payload["Desc"] = desc.strip()

    quantity = elem.findtext("Quantity")
    if quantity:
        payload["Quantity"] = Decimal(quantity.strip())

    uom = elem.findtext("UnitOfMeasure")
    if uom:
        payload["UnitOfMeasure"] = uom.strip()

    rate = elem.findtext("Rate")
    if rate:
        payload["Rate"] = Decimal(rate.strip())

    rate_percent = elem.findtext("RatePercent")
    if rate_percent:
        payload["RatePercent"] = Decimal(rate_percent.strip())

    price_level_ref = _parse_ref(elem.find("PriceLevelRef"))
    if price_level_ref is not None:
        payload["PriceLevelRef"] = price_level_ref

    class_ref = _parse_ref(elem.find("ClassRef"))
    if class_ref is not None:
        payload["ClassRef"] = class_ref

    amount = elem.findtext("Amount")
    if amount:
        payload["Amount"] = Decimal(amount.strip())

    inv_site = _parse_ref(elem.find("InventorySiteRef"))
    if inv_site is not None:
        payload["InventorySiteRef"] = inv_site

    inv_site_loc = _parse_ref(elem.find("InventorySiteLocationRef"))
    if inv_site_loc is not None:
        payload["InventorySiteLocationRef"] = inv_site_loc

    serial = elem.findtext("SerialNumber")
    if serial:
        payload["SerialNumber"] = serial.strip()

    lot = elem.findtext("LotNumber")
    if lot:
        payload["LotNumber"] = lot.strip()

    stc = _parse_ref(elem.find("SalesTaxCodeRef"))
    if stc is not None:
        payload["SalesTaxCodeRef"] = stc

    is_taxable = elem.findtext("IsTaxable")
    if is_taxable:
        payload["IsTaxable"] = is_taxable.strip().lower() == "true"

    is_manually_closed = elem.findtext("IsManuallyClosed")
    if is_manually_closed:
        payload["IsManuallyClosed"] = is_manually_closed.strip().lower() == "true"

    other1 = elem.findtext("Other1")
    if other1:
        payload["Other1"] = other1.strip()

    other2 = elem.findtext("Other2")
    if other2:
        payload["Other2"] = other2.strip()

    return SalesOrderLineItem.model_validate(payload)


def _parse_sales_order_ret(elem: etree._Element) -> SalesOrder:
    payload: dict[str, object] = {}

    for tag in (
        "TxnID",
        "TimeCreated",
        "TimeModified",
        "EditSequence",
        "TxnDate",
        "RefNumber",
        "PONumber",
        "DueDate",
        "FOB",
        "ShipDate",
        "Memo",
        "Other",
    ):
        text = elem.findtext(tag)
        if text:
            payload[tag] = text.strip()

    txn_number = elem.findtext("TxnNumber")
    if txn_number:
        payload["TxnNumber"] = int(txn_number.strip())

    # Refs
    for tag in (
        "CustomerRef",
        "ClassRef",
        "TemplateRef",
        "TermsRef",
        "SalesRepRef",
        "ShipMethodRef",
        "ItemSalesTaxRef",
        "CustomerSalesTaxCodeRef",
    ):
        ref_payload = _parse_ref(elem.find(tag))
        if ref_payload is not None:
            payload[tag] = ref_payload

    # Addresses
    bill = parse_address(elem.find("BillAddress"))
    if bill is not None:
        payload["BillAddress"] = bill.model_dump(by_alias=True, exclude_none=True)

    ship = parse_address(elem.find("ShipAddress"))
    if ship is not None:
        payload["ShipAddress"] = ship.model_dump(by_alias=True, exclude_none=True)

    # Booleans
    for tag in (
        "IsManuallyClosed",
        "IsToBePrinted",
        "IsToBeEmailed",
        "IsTaxIncluded",
        "IsFullyInvoiced",
    ):
        text = elem.findtext(tag)
        if text:
            payload[tag] = text.strip().lower() == "true"

    # Decimals
    for tag in (
        "Subtotal",
        "SalesTaxPercentage",
        "SalesTaxTotal",
        "TotalAmount",
        "ExchangeRate",
    ):
        text = elem.findtext(tag)
        if text:
            payload[tag] = Decimal(text.strip())

    # Line items
    line_items: list[SalesOrderLineItem] = []
    for line_elem in elem.findall("SalesOrderLineRet"):
        line_items.append(_parse_line_item(line_elem))
    payload["SalesOrderLineRet"] = line_items

    return SalesOrder.model_validate(payload)


def parse_query_response(xml: str) -> list[SalesOrder]:
    root = etree.fromstring(xml.encode("utf-8"))
    results: list[SalesOrder] = []
    for so_ret in root.iter("SalesOrderRet"):
        results.append(_parse_sales_order_ret(so_ret))
    return results


def parse_add_response(xml: str) -> SalesOrder:
    root = etree.fromstring(xml.encode("utf-8"))
    elem = root.find(".//SalesOrderRet")
    if elem is None:
        raise ValueError("No SalesOrderRet element found in Add response")
    return _parse_sales_order_ret(elem)


def parse_mod_response(xml: str) -> SalesOrder:
    root = etree.fromstring(xml.encode("utf-8"))
    elem = root.find(".//SalesOrderRet")
    if elem is None:
        raise ValueError("No SalesOrderRet element found in Mod response")
    return _parse_sales_order_ret(elem)
