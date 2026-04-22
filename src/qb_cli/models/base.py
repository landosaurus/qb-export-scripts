from __future__ import annotations

from pydantic import BaseModel, ConfigDict


class BaseEntity(BaseModel):
    model_config = ConfigDict(
        populate_by_name=True,
        extra="forbid",
        str_strip_whitespace=True,
    )
