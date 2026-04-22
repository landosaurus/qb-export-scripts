import pytest
from pydantic import Field, ValidationError

from qb_cli.models.base import BaseEntity


class Demo(BaseEntity):
    ref_number: str | None = Field(default=None, alias="RefNumber")
    memo: str | None = Field(default=None, alias="Memo")


def test_populate_by_name_and_alias() -> None:
    e1 = Demo(ref_number="1", memo="m")
    e2 = Demo.model_validate({"RefNumber": "2", "Memo": "n"})
    assert e1.ref_number == "1"
    assert e2.ref_number == "2"


def test_extra_fields_forbidden() -> None:
    with pytest.raises(ValidationError):
        Demo.model_validate({"RefNumber": "1", "NotARealField": "x"})


def test_dumps_with_qb_aliases_by_default() -> None:
    e = Demo(ref_number="1", memo="m")
    payload = e.model_dump(by_alias=True, exclude_none=True)
    assert payload == {"RefNumber": "1", "Memo": "m"}
