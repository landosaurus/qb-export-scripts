from __future__ import annotations

import csv
import sys
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Iterable, Type, TypeVar, Union, get_args, get_origin

from pydantic import BaseModel

from qb_cli.models.base import BaseEntity
from qb_cli.models.shared import Address, Ref

T = TypeVar("T", bound=BaseEntity)


# Column layout order for flattening an Address into CSV columns.
_ADDRESS_COLUMNS: tuple[str, ...] = (
    "addr1",
    "addr2",
    "addr3",
    "addr4",
    "addr5",
    "city",
    "state",
    "postal_code",
    "country",
    "note",
)


def _unwrap_optional(tp: object) -> object:
    """Strip Optional[...] / Union[X, None] wrappers and return the inner type.

    Returns the original annotation unchanged if it is not an Optional.
    """
    if get_origin(tp) is Union:
        args = tuple(a for a in get_args(tp) if a is not type(None))
        if len(args) == 1:
            return args[0]
    return tp


def _is_submodel_type(tp: object, parent: type[BaseModel]) -> bool:
    """True if ``tp`` is a BaseModel subclass other than the parent itself."""
    return isinstance(tp, type) and issubclass(tp, BaseModel) and tp is not parent


def _list_item_type(tp: object) -> object | None:
    """If ``tp`` is ``list[X]``, return X; else None."""
    if get_origin(tp) is list:
        args = get_args(tp)
        if args:
            return args[0]
    return None


def _classify_field(
    field_name: str, annotation: object
) -> tuple[str, list[str]]:
    """Classify a pydantic field and return (kind, column_names).

    kind is one of: "scalar", "ref", "address", "list", "skip".
    column_names is the list of CSV column names this field expands to
    (empty for "list" and "skip" — "list" is handled separately as line columns).
    """
    inner = _unwrap_optional(annotation)
    if inner is Ref:
        return "ref", [f"{field_name}_name"]
    if inner is Address:
        return "address", [f"{field_name}_{c}" for c in _ADDRESS_COLUMNS]
    if _list_item_type(inner) is not None:
        return "list", []
    # Scalars: str/int/Decimal/bool/date/datetime — anything else is skipped.
    if isinstance(inner, type) and issubclass(
        inner, (str, int, Decimal, bool, date, datetime)
    ):
        return "scalar", [field_name]
    # Unknown / unsupported type — skip with a warning.
    print(
        f"[csv_serializer] skipping field '{field_name}' with unsupported type {annotation!r}",
        file=sys.stderr,
    )
    return "skip", []


def _discover_columns(
    model: type[BaseModel],
) -> tuple[list[tuple[str, str, list[str]]], type[BaseModel] | None]:
    """Walk ``model.model_fields`` and return (column_specs, line_item_model).

    column_specs is a list of (field_name, kind, columns) for scalar/ref/address
    fields. line_item_model is the sub-model type for the (at most one) list
    field encountered, or None if there is no line-item field.
    """
    column_specs: list[tuple[str, str, list[str]]] = []
    line_item_model: type[BaseModel] | None = None
    for name, field in model.model_fields.items():
        kind, cols = _classify_field(name, field.annotation)
        if kind == "list":
            inner = _unwrap_optional(field.annotation)
            item_tp = _list_item_type(inner)
            if (
                isinstance(item_tp, type)
                and issubclass(item_tp, BaseModel)
                and line_item_model is None
            ):
                line_item_model = item_tp
            else:
                # Second list field or non-BaseModel list — not supported in v1.
                print(
                    f"[csv_serializer] skipping additional/unsupported list field '{name}'",
                    file=sys.stderr,
                )
            continue
        if kind == "skip":
            continue
        column_specs.append((name, kind, cols))
    return column_specs, line_item_model


def _render_scalar(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value)


def _fill_row_from_entity(
    entity: BaseModel,
    column_specs: list[tuple[str, str, list[str]]],
) -> dict[str, str]:
    """Serialize an entity's flat fields into a dict of CSV column -> string."""
    row: dict[str, str] = {}
    for name, kind, cols in column_specs:
        value = getattr(entity, name, None)
        if value is None:
            for c in cols:
                row[c] = ""
            continue
        if kind == "scalar":
            row[cols[0]] = _render_scalar(value)
        elif kind == "ref":
            assert isinstance(value, Ref)
            row[cols[0]] = value.full_name or ""
        elif kind == "address":
            assert isinstance(value, Address)
            for col_name, addr_attr in zip(cols, _ADDRESS_COLUMNS):
                row[col_name] = _render_scalar(getattr(value, addr_attr, None))
    return row


def _payload_from_row(
    row: dict[str, str],
    column_specs: list[tuple[str, str, list[str]]],
) -> dict[str, object]:
    """Un-flatten a CSV row into a dict suitable for ``model.model_validate``.

    Empty strings are dropped so pydantic sees absent fields as None.
    """
    payload: dict[str, object] = {}
    for name, kind, cols in column_specs:
        if kind == "scalar":
            raw = row.get(cols[0], "")
            if raw == "":
                continue
            payload[name] = raw
        elif kind == "ref":
            raw = row.get(cols[0], "")
            if raw == "":
                continue
            payload[name] = {"FullName": raw}
        elif kind == "address":
            addr_payload: dict[str, str] = {}
            for col_name, addr_attr in zip(cols, _ADDRESS_COLUMNS):
                raw = row.get(col_name, "")
                if raw == "":
                    continue
                # Use the pydantic alias (CamelCase) when feeding the payload
                # so Address can populate itself via alias.
                alias = Address.model_fields[addr_attr].alias or addr_attr
                addr_payload[alias] = raw
            if addr_payload:
                payload[name] = addr_payload
    return payload


def to_csv(entities: Iterable[BaseEntity], path: str | Path) -> None:
    entity_list = list(entities)
    # Use the first entity's class for column discovery; fall back to writing
    # just the meta columns if the input is empty (we can't infer a schema).
    if not entity_list:
        Path(path).write_text("row_type,parent_ref\n", encoding="utf-8")
        return

    model = type(entity_list[0])
    header_specs, line_model = _discover_columns(model)
    header_cols: list[str] = [c for _, _, cols in header_specs for c in cols]

    line_specs: list[tuple[str, str, list[str]]] = []
    line_cols: list[str] = []
    if line_model is not None:
        line_specs, _ = _discover_columns(line_model)
        line_cols = [c for _, _, cols in line_specs for c in cols]

    fieldnames: list[str] = ["row_type", "parent_ref"] + header_cols + line_cols

    with Path(path).open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, restval="")
        writer.writeheader()
        for entity in entity_list:
            header_row = _fill_row_from_entity(entity, header_specs)
            header_row["row_type"] = "header"
            header_row["parent_ref"] = ""
            writer.writerow(header_row)

            parent_ref: str = _render_scalar(getattr(entity, "ref_number", None))
            for item in getattr(entity, "line_items", []) or []:
                line_row = _fill_row_from_entity(item, line_specs)
                line_row["row_type"] = "line"
                line_row["parent_ref"] = parent_ref
                writer.writerow(line_row)


def from_csv(model: Type[T], path: str | Path) -> list[T]:
    header_specs, line_model = _discover_columns(model)
    line_specs: list[tuple[str, str, list[str]]] = []
    if line_model is not None:
        line_specs, _ = _discover_columns(line_model)

    header_rows: list[dict[str, str]] = []
    lines_by_parent: dict[str, list[dict[str, str]]] = {}

    with Path(path).open("r", newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            row_type = row.get("row_type", "")
            if row_type == "header":
                header_rows.append(row)
            elif row_type == "line":
                parent = row.get("parent_ref", "")
                lines_by_parent.setdefault(parent, []).append(row)
            # Silently ignore other row types for v1 forward-compat.

    result: list[T] = []
    for row in header_rows:
        payload = _payload_from_row(row, header_specs)
        ref_number = row.get("ref_number", "")
        if line_model is not None:
            line_payloads: list[dict[str, object]] = []
            for line_row in lines_by_parent.get(ref_number, []):
                line_payloads.append(_payload_from_row(line_row, line_specs))
            if line_payloads:
                payload["line_items"] = line_payloads
        result.append(model.model_validate(payload))

    return result
