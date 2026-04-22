from __future__ import annotations

import json
from pathlib import Path
from typing import Iterable, Type, TypeVar

from qb_cli.models.base import BaseEntity

T = TypeVar("T", bound=BaseEntity)


def to_json(entities: Iterable[BaseEntity], path: str | Path) -> None:
    records = [e.model_dump(by_alias=True, exclude_none=True, mode="json") for e in entities]
    Path(path).write_text(json.dumps(records, indent=2))


def from_json(model: Type[T], path: str | Path) -> list[T]:
    raw = json.loads(Path(path).read_text())
    return [model.model_validate(r) for r in raw]
