from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from typing import Optional

from pydantic import Field

from qb_cli.models.base import BaseEntity


def quantize_money(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


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
