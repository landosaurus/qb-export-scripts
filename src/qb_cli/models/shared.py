from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from typing import Annotated, Optional

from pydantic import BeforeValidator, Field

from qb_cli.models.base import BaseEntity


def quantize_money(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _coerce_money(v: object) -> Optional[Decimal]:
    """Pydantic BeforeValidator that ensures every money value has exactly 2dp.

    QuickBooks rejects Amount values with other precision (e.g. ``97.5`` fails
    as ``Amount: There was an error when converting the amount "97.5"``). Since
    Excel silently strips trailing zeros when a user edits an exported CSV,
    we quantize on the way in so the model always carries canonical 2dp values.
    """
    if v is None or v == "":
        return None
    if isinstance(v, Decimal):
        return quantize_money(v)
    return quantize_money(Decimal(str(v)))


Money = Annotated[Optional[Decimal], BeforeValidator(_coerce_money)]


class Address(BaseEntity):
    addr1: Optional[str] = Field(default=None, alias="Addr1")
    addr2: Optional[str] = Field(default=None, alias="Addr2")
    addr3: Optional[str] = Field(default=None, alias="Addr3")
    addr4: Optional[str] = Field(default=None, alias="Addr4")
    addr5: Optional[str] = Field(default=None, alias="Addr5")
    city: Optional[str] = Field(default=None, alias="City")
    state: Optional[str] = Field(default=None, alias="State")
    postal_code: Optional[str] = Field(default=None, alias="PostalCode")
    country: Optional[str] = Field(default=None, alias="Country")
    note: Optional[str] = Field(default=None, alias="Note")

    def render_multiline(self) -> str:
        lines: list[str] = []
        for line in (self.addr1, self.addr2, self.addr3, self.addr4, self.addr5):
            if line:
                lines.append(line)
        city_line = ", ".join(p for p in (self.city, self.state) if p)
        if city_line and self.postal_code:
            city_line = f"{city_line} {self.postal_code}"
        elif self.postal_code:
            city_line = self.postal_code
        if city_line:
            lines.append(city_line)
        if self.country:
            lines.append(self.country)
        return "\n".join(lines)


class Ref(BaseEntity):
    list_id: Optional[str] = Field(default=None, alias="ListID")
    full_name: Optional[str] = Field(default=None, alias="FullName")
