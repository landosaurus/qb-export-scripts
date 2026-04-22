from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from typing import ClassVar, Optional

from pydantic import Field

from qb_cli.models.base import BaseEntity
from qb_cli.models.shared import Address, Ref


class InvoiceLineItem(BaseEntity):
    txn_line_id: Optional[str] = Field(default=None, alias="TxnLineID")
    item_ref: Optional[Ref] = Field(default=None, alias="ItemRef")
    override_item_account_ref: Optional[Ref] = Field(
        default=None, alias="OverrideItemAccountRef"
    )
    description: Optional[str] = Field(default=None, alias="Desc", max_length=4095)
    quantity: Optional[Decimal] = Field(default=None, alias="Quantity")
    unit_of_measure: Optional[str] = Field(default=None, alias="UnitOfMeasure")
    rate: Optional[Decimal] = Field(default=None, alias="Rate")
    rate_percent: Optional[Decimal] = Field(default=None, alias="RatePercent")
    price_level_ref: Optional[Ref] = Field(default=None, alias="PriceLevelRef")
    class_ref: Optional[Ref] = Field(default=None, alias="ClassRef")
    amount: Optional[Decimal] = Field(default=None, alias="Amount")
    inventory_site_ref: Optional[Ref] = Field(default=None, alias="InventorySiteRef")
    inventory_site_location_ref: Optional[Ref] = Field(
        default=None, alias="InventorySiteLocationRef"
    )
    serial_number: Optional[str] = Field(default=None, alias="SerialNumber", max_length=255)
    lot_number: Optional[str] = Field(default=None, alias="LotNumber", max_length=255)
    service_date: Optional[date] = Field(default=None, alias="ServiceDate")
    sales_tax_code_ref: Optional[Ref] = Field(default=None, alias="SalesTaxCodeRef")
    is_taxable: Optional[bool] = Field(default=None, alias="IsTaxable")
    customer_ref: Optional[Ref] = Field(default=None, alias="CustomerRef")
    other1: Optional[str] = Field(default=None, alias="Other1", max_length=255)
    other2: Optional[str] = Field(default=None, alias="Other2", max_length=255)


class Invoice(BaseEntity):
    REQUIRED_ON_ADD: ClassVar[list[str]] = ["customer_ref"]

    # Identity / read-only
    txn_id: Optional[str] = Field(default=None, alias="TxnID")
    time_created: Optional[datetime] = Field(default=None, alias="TimeCreated")
    time_modified: Optional[datetime] = Field(default=None, alias="TimeModified")
    edit_sequence: Optional[str] = Field(default=None, alias="EditSequence")
    txn_number: Optional[int] = Field(default=None, alias="TxnNumber")

    # Header
    customer_ref: Optional[Ref] = Field(default=None, alias="CustomerRef")
    class_ref: Optional[Ref] = Field(default=None, alias="ClassRef")
    ar_account_ref: Optional[Ref] = Field(default=None, alias="ARAccountRef")
    template_ref: Optional[Ref] = Field(default=None, alias="TemplateRef")
    txn_date: Optional[date] = Field(default=None, alias="TxnDate")
    ref_number: Optional[str] = Field(default=None, alias="RefNumber", max_length=11)
    bill_address: Optional[Address] = Field(default=None, alias="BillAddress")
    ship_address: Optional[Address] = Field(default=None, alias="ShipAddress")
    is_pending: Optional[bool] = Field(default=None, alias="IsPending")
    po_number: Optional[str] = Field(default=None, alias="PONumber", max_length=25)
    terms_ref: Optional[Ref] = Field(default=None, alias="TermsRef")
    due_date: Optional[date] = Field(default=None, alias="DueDate")
    sales_rep_ref: Optional[Ref] = Field(default=None, alias="SalesRepRef")
    fob: Optional[str] = Field(default=None, alias="FOB", max_length=13)
    ship_date: Optional[date] = Field(default=None, alias="ShipDate")
    ship_method_ref: Optional[Ref] = Field(default=None, alias="ShipMethodRef")
    item_sales_tax_ref: Optional[Ref] = Field(default=None, alias="ItemSalesTaxRef")
    memo: Optional[str] = Field(default=None, alias="Memo", max_length=4095)
    customer_msg_ref: Optional[Ref] = Field(default=None, alias="CustomerMsgRef")
    is_to_be_printed: Optional[bool] = Field(default=None, alias="IsToBePrinted")
    is_to_be_emailed: Optional[bool] = Field(default=None, alias="IsToBeEmailed")
    is_tax_included: Optional[bool] = Field(default=None, alias="IsTaxIncluded")
    customer_sales_tax_code_ref: Optional[Ref] = Field(
        default=None, alias="CustomerSalesTaxCodeRef"
    )
    other: Optional[str] = Field(default=None, alias="Other", max_length=29)
    exchange_rate: Optional[Decimal] = Field(default=None, alias="ExchangeRate")

    # Line items
    line_items: list[InvoiceLineItem] = Field(
        default_factory=list, alias="InvoiceLineRet"
    )

    # Totals (read-only)
    subtotal: Optional[Decimal] = Field(default=None, alias="Subtotal")
    sales_tax_percentage: Optional[Decimal] = Field(
        default=None, alias="SalesTaxPercentage"
    )
    sales_tax_total: Optional[Decimal] = Field(default=None, alias="SalesTaxTotal")
    applied_amount: Optional[Decimal] = Field(default=None, alias="AppliedAmount")
    balance_remaining: Optional[Decimal] = Field(default=None, alias="BalanceRemaining")
    total_amount: Optional[Decimal] = Field(default=None, alias="TotalAmount")
    is_paid: Optional[bool] = Field(default=None, alias="IsPaid")
    is_finance_charge: Optional[bool] = Field(default=None, alias="IsFinanceCharge")
    suggested_discount_amount: Optional[Decimal] = Field(
        default=None, alias="SuggestedDiscountAmount"
    )
