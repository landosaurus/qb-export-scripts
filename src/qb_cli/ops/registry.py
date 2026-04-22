from __future__ import annotations

from dataclasses import dataclass
from types import ModuleType
from typing import Type

from qb_cli.models.base import BaseEntity
from qb_cli.models.invoice import Invoice
from qb_cli.models.purchase_order import PurchaseOrder
from qb_cli.models.sales_order import SalesOrder
from qb_cli.qbxml import invoice as invoice_qbxml
from qb_cli.qbxml import purchase_order as purchase_order_qbxml
from qb_cli.qbxml import sales_order as sales_order_qbxml


@dataclass(frozen=True)
class EntityHandler:
    key: str
    model: Type[BaseEntity]
    qbxml: ModuleType
    id_field: str
    supports_mod: bool = True
    has_line_items: bool = True


HANDLERS: dict[str, EntityHandler] = {
    "invoice": EntityHandler("invoice", Invoice, invoice_qbxml, id_field="ref_number"),
    "sales_order": EntityHandler(
        "sales_order", SalesOrder, sales_order_qbxml, id_field="ref_number"
    ),
    "purchase_order": EntityHandler(
        "purchase_order", PurchaseOrder, purchase_order_qbxml, id_field="ref_number"
    ),
}


def get_handler(key: str) -> EntityHandler:
    try:
        return HANDLERS[key]
    except KeyError:
        raise KeyError(
            f"unknown entity key '{key}' — supported: {', '.join(sorted(HANDLERS))}"
        )
