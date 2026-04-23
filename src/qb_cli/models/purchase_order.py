from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import ClassVar, Optional

from pydantic import Field

from qb_cli.models.base import BaseEntity
from qb_cli.models.shared import Address, Money, Ref


class PurchaseOrderLineItem(BaseEntity):
    txn_line_id: Optional[str] = Field(default=None, alias="TxnLineID")
    item_ref: Optional[Ref] = Field(default=None, alias="ItemRef")
    manufacturer_part_number: Optional[str] = Field(
        default=None, alias="ManufacturerPartNumber", max_length=31
    )
    description: Optional[str] = Field(default=None, alias="Desc", max_length=4095)
    quantity: Optional[Decimal] = Field(default=None, alias="Quantity")
    unit_of_measure: Optional[str] = Field(default=None, alias="UnitOfMeasure")
    rate: Optional[Decimal] = Field(default=None, alias="Rate")
    class_ref: Optional[Ref] = Field(default=None, alias="ClassRef")
    amount: Money = Field(default=None, alias="Amount")
    inventory_site_ref: Optional[Ref] = Field(default=None, alias="InventorySiteRef")
    inventory_site_location_ref: Optional[Ref] = Field(
        default=None, alias="InventorySiteLocationRef"
    )
    customer_ref: Optional[Ref] = Field(default=None, alias="CustomerRef")
    service_date: Optional[date] = Field(default=None, alias="ServiceDate")
    sales_tax_code_ref: Optional[Ref] = Field(default=None, alias="SalesTaxCodeRef")
    received_quantity: Optional[Decimal] = Field(default=None, alias="ReceivedQuantity")
    is_manually_closed: Optional[bool] = Field(default=None, alias="IsManuallyClosed")
    other1: Optional[str] = Field(default=None, alias="Other1")
    other2: Optional[str] = Field(default=None, alias="Other2")


class PurchaseOrder(BaseEntity):
    REQUIRED_ON_ADD: ClassVar[list[str]] = ["vendor_ref"]

    # Identity / read-only
    txn_id: Optional[str] = Field(default=None, alias="TxnID")
    time_created: Optional[str] = Field(default=None, alias="TimeCreated")
    time_modified: Optional[str] = Field(default=None, alias="TimeModified")
    edit_sequence: Optional[str] = Field(default=None, alias="EditSequence")
    txn_number: Optional[int] = Field(default=None, alias="TxnNumber")

    # Header
    vendor_ref: Optional[Ref] = Field(default=None, alias="VendorRef")
    class_ref: Optional[Ref] = Field(default=None, alias="ClassRef")
    template_ref: Optional[Ref] = Field(default=None, alias="TemplateRef")
    txn_date: Optional[date] = Field(default=None, alias="TxnDate")
    ref_number: Optional[str] = Field(default=None, alias="RefNumber", max_length=11)
    vendor_address: Optional[Address] = Field(default=None, alias="VendorAddress")
    ship_address: Optional[Address] = Field(default=None, alias="ShipAddress")
    terms_ref: Optional[Ref] = Field(default=None, alias="TermsRef")
    due_date: Optional[date] = Field(default=None, alias="DueDate")
    expected_date: Optional[date] = Field(default=None, alias="ExpectedDate")
    ship_method_ref: Optional[Ref] = Field(default=None, alias="ShipMethodRef")
    fob: Optional[str] = Field(default=None, alias="FOB", max_length=13)
    memo: Optional[str] = Field(default=None, alias="Memo", max_length=4095)
    vendor_msg: Optional[str] = Field(default=None, alias="VendorMsg", max_length=99)
    is_to_be_printed: Optional[bool] = Field(default=None, alias="IsToBePrinted")
    is_to_be_emailed: Optional[bool] = Field(default=None, alias="IsToBeEmailed")
    is_tax_included: Optional[bool] = Field(default=None, alias="IsTaxIncluded")
    sales_tax_code_ref: Optional[Ref] = Field(default=None, alias="SalesTaxCodeRef")
    other1: Optional[str] = Field(default=None, alias="Other1", max_length=29)
    other2: Optional[str] = Field(default=None, alias="Other2", max_length=29)
    exchange_rate: Optional[Decimal] = Field(default=None, alias="ExchangeRate")
    is_manually_closed: Optional[bool] = Field(default=None, alias="IsManuallyClosed")

    # Read-only (from query)
    subtotal: Money = Field(default=None, alias="Subtotal")
    sales_tax_total: Money = Field(default=None, alias="SalesTaxTotal")
    total_amount: Money = Field(default=None, alias="TotalAmount")
    is_fully_received: Optional[bool] = Field(default=None, alias="IsFullyReceived")

    # Line items
    line_items: list[PurchaseOrderLineItem] = Field(
        default_factory=list, alias="PurchaseOrderLineRet"
    )
