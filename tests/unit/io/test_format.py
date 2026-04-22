import pytest
from qb_cli.io.format import Format, detect_format


def test_detect_by_extension():
    assert detect_format("x.csv") is Format.CSV
    assert detect_format("x.json") is Format.JSON
    assert detect_format("/path/y.JSON") is Format.JSON


def test_detect_unknown_raises():
    with pytest.raises(ValueError):
        detect_format("x.yaml")
