import pytest
from qb_cli.transport.errors import (
    QBError,
    QBConnectionError,
    QBXMLParseError,
    QBStatusError,
    QBEntityNotFound,
    QBDuplicateEntity,
)


def test_qb_status_error_carries_fields():
    err = QBStatusError(
        status_code=3100,
        status_message="name already exists",
        status_severity="Error",
        request_id="1",
    )
    assert err.status_code == 3100
    assert "3100" in str(err)
    assert "name already exists" in str(err)


def test_entity_not_found_is_status_error():
    err = QBEntityNotFound(status_code=500, status_message="not found", status_severity="Error", request_id="1")
    assert isinstance(err, QBStatusError)
    assert isinstance(err, QBError)


def test_duplicate_entity_is_status_error():
    err = QBDuplicateEntity(status_code=3100, status_message="dup", status_severity="Error", request_id="1")
    assert isinstance(err, QBStatusError)


def test_base_error_hierarchy():
    assert issubclass(QBConnectionError, QBError)
    assert issubclass(QBXMLParseError, QBError)
    assert issubclass(QBStatusError, QBError)
