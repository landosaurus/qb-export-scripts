from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import Optional

from lxml import etree

from qb_cli.models.invoice import Invoice, InvoiceLineItem
from qb_cli.models.shared import Address, Ref
from qb_cli.qbxml.common import format_qb_date, parse_address, xml_escape
from qb_cli.qbxml.envelope import wrap_request


# ---------------------------------------------------------------------------
# Builders
# ---------------------------------------------------------------------------


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
    """Build an InvoiceQueryRq request body wrapped in a QBXML envelope."""
    # Iterator attributes
    iterator_attrs = ""
    if iterator_id is not None:
        iterator_attrs = f' iterator="Continue" iteratorID="{xml_escape(iterator_id)}"'
    elif max_returned is not None and not txn_ids and not ref_numbers:
        iterator_attrs = ' iterator="Start"'

    parts: list[str] = [f'    <InvoiceQueryRq requestID="1"{iterator_attrs}>']

    # 1. TxnID filters
    if txn_ids:
        for tid in txn_ids:
            parts.append(f'      <TxnID>{xml_escape(tid)}</TxnID>')

    # 2. RefNumber filters
    if ref_numbers:
        for ref in ref_numbers:
            parts.append(f'      <RefNumber>{xml_escape(ref)}</RefNumber>')

    # 3. MaxReturned — forbidden when filtering by TxnID/RefNumber
    if max_returned is not None and not txn_ids and not ref_numbers:
        parts.append(f'      <MaxReturned>{max_returned}</MaxReturned>')

    # 4. TxnDateRangeFilter (MUST come before IncludeLineItems)
    if date_from is not None or date_to is not None:
        parts.append('      <TxnDateRangeFilter>')
        if date_from is not None:
            parts.append(f'        <FromTxnDate>{format_qb_date(date_from)}</FromTxnDate>')
        if date_to is not None:
            parts.append(f'        <ToTxnDate>{format_qb_date(date_to)}</ToTxnDate>')
        parts.append('      </TxnDateRangeFilter>')

    # 5. IncludeLineItems
    if include_line_items:
        parts.append('      <IncludeLineItems>true</IncludeLineItems>')

    parts.append('    </InvoiceQueryRq>')

    return wrap_request("\n".join(parts))


def _emit_ref(ref: Ref, tag: str, indent: str) -> list[str]:
    """Emit a reference element (ListID preferred, then FullName)."""
    lines = [f'{indent}<{tag}>']
    if ref.list_id is not None:
        lines.append(f'{indent}  <ListID>{xml_escape(ref.list_id)}</ListID>')
    elif ref.full_name is not None:
        lines.append(f'{indent}  <FullName>{xml_escape(ref.full_name)}</FullName>')
    lines.append(f'{indent}</{tag}>')
    return lines


def _emit_address(addr: Address, tag: str, indent: str) -> list[str]:
    """Emit an address block."""
    lines = [f'{indent}<{tag}>']
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
        value: Optional[str] = getattr(addr, field_name)
        if value is not None:
            lines.append(f'{indent}  <{alias}>{xml_escape(value)}</{alias}>')
    lines.append(f'{indent}</{tag}>')
    return lines


def _emit_decimal(value: Decimal, tag: str, indent: str) -> str:
    return f'{indent}<{tag}>{str(value)}</{tag}>'


def _emit_bool(value: bool, tag: str, indent: str) -> str:
    return f'{indent}<{tag}>{"true" if value else "false"}</{tag}>'


def _build_line_item_block(
    line: InvoiceLineItem,
    wrapper_tag: str,
    indent: str,
) -> list[str]:
    """Build an InvoiceLineAdd or InvoiceLineMod element."""
    inner = indent + "  "
    lines: list[str] = [f'{indent}<{wrapper_tag}>']

    # TxnLineID — required for Mod; for Add only emit if present
    if line.txn_line_id is not None:
        lines.append(f'{inner}<TxnLineID>{xml_escape(line.txn_line_id)}</TxnLineID>')
    elif wrapper_tag == "InvoiceLineMod":
        # Mod treats missing TxnLineID as a new line
        lines.append(f'{inner}<TxnLineID>-1</TxnLineID>')

    if line.item_ref is not None:
        lines.extend(_emit_ref(line.item_ref, "ItemRef", inner))

    if line.override_item_account_ref is not None:
        lines.extend(
            _emit_ref(line.override_item_account_ref, "OverrideItemAccountRef", inner)
        )

    if line.description is not None:
        lines.append(f'{inner}<Desc>{xml_escape(line.description)}</Desc>')

    if line.quantity is not None:
        lines.append(_emit_decimal(line.quantity, "Quantity", inner))

    if line.unit_of_measure is not None:
        lines.append(
            f'{inner}<UnitOfMeasure>{xml_escape(line.unit_of_measure)}</UnitOfMeasure>'
        )

    if line.rate is not None:
        lines.append(_emit_decimal(line.rate, "Rate", inner))
    elif line.rate_percent is not None:
        lines.append(_emit_decimal(line.rate_percent, "RatePercent", inner))

    if line.price_level_ref is not None:
        lines.extend(_emit_ref(line.price_level_ref, "PriceLevelRef", inner))

    if line.class_ref is not None:
        lines.extend(_emit_ref(line.class_ref, "ClassRef", inner))

    if line.amount is not None:
        lines.append(_emit_decimal(line.amount, "Amount", inner))

    if line.inventory_site_ref is not None:
        lines.extend(_emit_ref(line.inventory_site_ref, "InventorySiteRef", inner))

    if line.inventory_site_location_ref is not None:
        lines.extend(
            _emit_ref(
                line.inventory_site_location_ref, "InventorySiteLocationRef", inner
            )
        )

    if line.serial_number is not None:
        lines.append(
            f'{inner}<SerialNumber>{xml_escape(line.serial_number)}</SerialNumber>'
        )

    if line.lot_number is not None:
        lines.append(f'{inner}<LotNumber>{xml_escape(line.lot_number)}</LotNumber>')

    if line.service_date is not None:
        lines.append(
            f'{inner}<ServiceDate>{format_qb_date(line.service_date)}</ServiceDate>'
        )

    if line.sales_tax_code_ref is not None:
        lines.extend(_emit_ref(line.sales_tax_code_ref, "SalesTaxCodeRef", inner))

    if line.is_taxable is not None:
        lines.append(_emit_bool(line.is_taxable, "IsTaxable", inner))

    if line.customer_ref is not None:
        lines.extend(_emit_ref(line.customer_ref, "CustomerRef", inner))

    if line.other1 is not None:
        lines.append(f'{inner}<Other1>{xml_escape(line.other1)}</Other1>')
    if line.other2 is not None:
        lines.append(f'{inner}<Other2>{xml_escape(line.other2)}</Other2>')

    lines.append(f'{indent}</{wrapper_tag}>')
    return lines


def _build_invoice_body(
    entity: Invoice,
    *,
    include_identity: bool,
    line_wrapper_tag: str,
    indent: str,
) -> list[str]:
    """Emit the inner body (children of InvoiceAdd or InvoiceMod)."""
    inner = indent
    lines: list[str] = []

    if include_identity:
        # TxnID & EditSequence come first for Mod
        if entity.txn_id is not None:
            lines.append(f'{inner}<TxnID>{xml_escape(entity.txn_id)}</TxnID>')
        if entity.edit_sequence is not None:
            lines.append(
                f'{inner}<EditSequence>{xml_escape(entity.edit_sequence)}</EditSequence>'
            )

    # CustomerRef
    if entity.customer_ref is not None:
        lines.extend(_emit_ref(entity.customer_ref, "CustomerRef", inner))

    # ClassRef
    if entity.class_ref is not None:
        lines.extend(_emit_ref(entity.class_ref, "ClassRef", inner))

    # ARAccountRef
    if entity.ar_account_ref is not None:
        lines.extend(_emit_ref(entity.ar_account_ref, "ARAccountRef", inner))

    # TemplateRef (before TxnDate per schema)
    if entity.template_ref is not None:
        lines.extend(_emit_ref(entity.template_ref, "TemplateRef", inner))

    # TxnDate
    if entity.txn_date is not None:
        lines.append(f'{inner}<TxnDate>{format_qb_date(entity.txn_date)}</TxnDate>')

    # RefNumber
    if entity.ref_number is not None:
        lines.append(f'{inner}<RefNumber>{xml_escape(entity.ref_number)}</RefNumber>')

    # BillAddress
    if entity.bill_address is not None:
        lines.extend(_emit_address(entity.bill_address, "BillAddress", inner))

    # ShipAddress
    if entity.ship_address is not None:
        lines.extend(_emit_address(entity.ship_address, "ShipAddress", inner))

    # IsPending
    if entity.is_pending is not None:
        lines.append(_emit_bool(entity.is_pending, "IsPending", inner))

    # PONumber
    if entity.po_number is not None:
        lines.append(f'{inner}<PONumber>{xml_escape(entity.po_number)}</PONumber>')

    # TermsRef
    if entity.terms_ref is not None:
        lines.extend(_emit_ref(entity.terms_ref, "TermsRef", inner))

    # DueDate
    if entity.due_date is not None:
        lines.append(f'{inner}<DueDate>{format_qb_date(entity.due_date)}</DueDate>')

    # SalesRepRef
    if entity.sales_rep_ref is not None:
        lines.extend(_emit_ref(entity.sales_rep_ref, "SalesRepRef", inner))

    # FOB
    if entity.fob is not None:
        lines.append(f'{inner}<FOB>{xml_escape(entity.fob)}</FOB>')

    # ShipDate
    if entity.ship_date is not None:
        lines.append(f'{inner}<ShipDate>{format_qb_date(entity.ship_date)}</ShipDate>')

    # ShipMethodRef
    if entity.ship_method_ref is not None:
        lines.extend(_emit_ref(entity.ship_method_ref, "ShipMethodRef", inner))

    # ItemSalesTaxRef
    if entity.item_sales_tax_ref is not None:
        lines.extend(_emit_ref(entity.item_sales_tax_ref, "ItemSalesTaxRef", inner))

    # Memo
    if entity.memo is not None:
        lines.append(f'{inner}<Memo>{xml_escape(entity.memo)}</Memo>')

    # CustomerMsgRef
    if entity.customer_msg_ref is not None:
        lines.extend(_emit_ref(entity.customer_msg_ref, "CustomerMsgRef", inner))

    # IsToBePrinted
    if entity.is_to_be_printed is not None:
        lines.append(_emit_bool(entity.is_to_be_printed, "IsToBePrinted", inner))

    # IsToBeEmailed
    if entity.is_to_be_emailed is not None:
        lines.append(_emit_bool(entity.is_to_be_emailed, "IsToBeEmailed", inner))

    # IsTaxIncluded
    if entity.is_tax_included is not None:
        lines.append(_emit_bool(entity.is_tax_included, "IsTaxIncluded", inner))

    # CustomerSalesTaxCodeRef
    if entity.customer_sales_tax_code_ref is not None:
        lines.extend(
            _emit_ref(
                entity.customer_sales_tax_code_ref, "CustomerSalesTaxCodeRef", inner
            )
        )

    # Other
    if entity.other is not None:
        lines.append(f'{inner}<Other>{xml_escape(entity.other)}</Other>')

    # ExchangeRate
    if entity.exchange_rate is not None:
        lines.append(_emit_decimal(entity.exchange_rate, "ExchangeRate", inner))

    # Line items
    for line in entity.line_items:
        lines.extend(_build_line_item_block(line, line_wrapper_tag, inner))

    return lines


def build_add(entity: Invoice) -> str:
    """Build an InvoiceAddRq request."""
    body_lines = _build_invoice_body(
        entity,
        include_identity=False,
        line_wrapper_tag="InvoiceLineAdd",
        indent="        ",
    )
    parts: list[str] = [
        '    <InvoiceAddRq>',
        '      <InvoiceAdd>',
    ]
    parts.extend(body_lines)
    parts.extend(
        [
            '      </InvoiceAdd>',
            '    </InvoiceAddRq>',
        ]
    )
    return wrap_request("\n".join(parts))


def build_mod(entity: Invoice) -> str:
    """Build an InvoiceModRq request. Requires edit_sequence and txn_id."""
    if entity.edit_sequence is None:
        raise ValueError("build_mod requires EditSequence")
    if entity.txn_id is None:
        raise ValueError("build_mod requires TxnID")

    body_lines = _build_invoice_body(
        entity,
        include_identity=True,
        line_wrapper_tag="InvoiceLineMod",
        indent="        ",
    )
    parts: list[str] = [
        '    <InvoiceModRq>',
        '      <InvoiceMod>',
    ]
    parts.extend(body_lines)
    parts.extend(
        [
            '      </InvoiceMod>',
            '    </InvoiceModRq>',
        ]
    )
    return wrap_request("\n".join(parts))


# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------


def _text_or_none(elem: Optional[etree._Element]) -> Optional[str]:
    if elem is None:
        return None
    text = elem.text
    if text is None:
        return None
    stripped = text.strip()
    if not stripped:
        return None
    return stripped


def _parse_ref(elem: Optional[etree._Element]) -> Optional[dict[str, str]]:
    if elem is None:
        return None
    payload: dict[str, str] = {}
    list_id = _text_or_none(elem.find("ListID"))
    full_name = _text_or_none(elem.find("FullName"))
    if list_id is not None:
        payload["ListID"] = list_id
    if full_name is not None:
        payload["FullName"] = full_name
    if not payload:
        return None
    return payload


def _parse_address_dict(elem: Optional[etree._Element]) -> Optional[dict[str, str]]:
    addr = parse_address(elem)
    if addr is None:
        return None
    dumped = addr.model_dump(by_alias=True, exclude_none=True)
    # Cast values to str explicitly — Address fields are all Optional[str]
    return {str(k): str(v) for k, v in dumped.items()}


_LINE_REF_FIELDS: tuple[str, ...] = (
    "ItemRef",
    "OverrideItemAccountRef",
    "PriceLevelRef",
    "ClassRef",
    "InventorySiteRef",
    "InventorySiteLocationRef",
    "SalesTaxCodeRef",
    "CustomerRef",
)

_LINE_TEXT_FIELDS: tuple[str, ...] = (
    "TxnLineID",
    "Desc",
    "Quantity",
    "UnitOfMeasure",
    "Rate",
    "RatePercent",
    "Amount",
    "SerialNumber",
    "LotNumber",
    "ServiceDate",
    "Other1",
    "Other2",
)


def _parse_line_item(elem: etree._Element) -> dict[str, object]:
    payload: dict[str, object] = {}
    for tag in _LINE_TEXT_FIELDS:
        value = _text_or_none(elem.find(tag))
        if value is not None:
            payload[tag] = value
    for tag in _LINE_REF_FIELDS:
        ref = _parse_ref(elem.find(tag))
        if ref is not None:
            payload[tag] = ref
    is_taxable = _text_or_none(elem.find("IsTaxable"))
    if is_taxable is not None:
        payload["IsTaxable"] = is_taxable == "true"
    return payload


_HEADER_REF_FIELDS: tuple[str, ...] = (
    "CustomerRef",
    "ClassRef",
    "ARAccountRef",
    "TemplateRef",
    "TermsRef",
    "SalesRepRef",
    "ShipMethodRef",
    "ItemSalesTaxRef",
    "CustomerMsgRef",
    "CustomerSalesTaxCodeRef",
)

_HEADER_TEXT_FIELDS: tuple[str, ...] = (
    "TxnID",
    "TimeCreated",
    "TimeModified",
    "EditSequence",
    "TxnNumber",
    "TxnDate",
    "RefNumber",
    "PONumber",
    "DueDate",
    "FOB",
    "ShipDate",
    "Memo",
    "Other",
    "ExchangeRate",
    "Subtotal",
    "SalesTaxPercentage",
    "SalesTaxTotal",
    "AppliedAmount",
    "BalanceRemaining",
    "TotalAmount",
    "SuggestedDiscountAmount",
)

_HEADER_BOOL_FIELDS: tuple[str, ...] = (
    "IsPending",
    "IsToBePrinted",
    "IsToBeEmailed",
    "IsTaxIncluded",
    "IsPaid",
    "IsFinanceCharge",
)


def _parse_invoice_ret(elem: etree._Element) -> Invoice:
    payload: dict[str, object] = {}

    for tag in _HEADER_TEXT_FIELDS:
        value = _text_or_none(elem.find(tag))
        if value is not None:
            payload[tag] = value

    for tag in _HEADER_BOOL_FIELDS:
        value = _text_or_none(elem.find(tag))
        if value is not None:
            payload[tag] = value == "true"

    for tag in _HEADER_REF_FIELDS:
        ref = _parse_ref(elem.find(tag))
        if ref is not None:
            payload[tag] = ref

    # Addresses
    bill_addr = _parse_address_dict(elem.find("BillAddress"))
    if bill_addr is not None:
        payload["BillAddress"] = bill_addr

    ship_addr = _parse_address_dict(elem.find("ShipAddress"))
    if ship_addr is not None:
        payload["ShipAddress"] = ship_addr

    # Line items
    line_payloads: list[dict[str, object]] = []
    for line_elem in elem.findall("InvoiceLineRet"):
        line_payloads.append(_parse_line_item(line_elem))
    payload["InvoiceLineRet"] = line_payloads

    return Invoice.model_validate(payload)


def parse_query_response(xml: str) -> list[Invoice]:
    """Parse an InvoiceQueryRs response into a list of Invoice models."""
    root = etree.fromstring(xml.encode("utf-8"))
    results: list[Invoice] = []
    for elem in root.iter("InvoiceRet"):
        results.append(_parse_invoice_ret(elem))
    return results


def _parse_single_ret(xml: str) -> Invoice:
    root = etree.fromstring(xml.encode("utf-8"))
    elem = root.find(".//InvoiceRet")
    if elem is None:
        raise ValueError("No InvoiceRet element found in response")
    return _parse_invoice_ret(elem)


def parse_add_response(xml: str) -> Invoice:
    """Parse an InvoiceAddRs response into an Invoice model."""
    return _parse_single_ret(xml)


def parse_mod_response(xml: str) -> Invoice:
    """Parse an InvoiceModRs response into an Invoice model."""
    return _parse_single_ret(xml)
