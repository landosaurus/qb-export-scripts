import pytest
from qb_cli.transport.status import check_response_status
from qb_cli.transport.errors import QBEntityNotFound, QBDuplicateEntity, QBStatusError


OK_XML = """<?xml version="1.0"?>
<QBXML><QBXMLMsgsRs>
  <InvoiceQueryRs requestID="1" statusCode="0" statusSeverity="Info" statusMessage="Status OK"/>
</QBXMLMsgsRs></QBXML>"""

NOT_FOUND_XML = """<?xml version="1.0"?>
<QBXML><QBXMLMsgsRs>
  <InvoiceQueryRs requestID="1" statusCode="500" statusSeverity="Error" statusMessage="No matching"/>
</QBXMLMsgsRs></QBXML>"""

DUP_XML = """<?xml version="1.0"?>
<QBXML><QBXMLMsgsRs>
  <InvoiceAddRs requestID="1" statusCode="3100" statusSeverity="Error" statusMessage="name already exists"/>
</QBXMLMsgsRs></QBXML>"""

OTHER_ERR_XML = """<?xml version="1.0"?>
<QBXML><QBXMLMsgsRs>
  <InvoiceAddRs requestID="1" statusCode="1" statusSeverity="Warn" statusMessage="something"/>
</QBXMLMsgsRs></QBXML>"""


def test_ok_returns_normally():
    check_response_status(OK_XML)


def test_not_found_raises_entity_not_found():
    with pytest.raises(QBEntityNotFound) as ei:
        check_response_status(NOT_FOUND_XML)
    assert ei.value.status_code == 500


def test_duplicate_raises_duplicate_entity():
    with pytest.raises(QBDuplicateEntity) as ei:
        check_response_status(DUP_XML)
    assert ei.value.status_code == 3100


def test_other_error_raises_plain_status_error():
    with pytest.raises(QBStatusError) as ei:
        check_response_status(OTHER_ERR_XML)
    assert not isinstance(ei.value, QBDuplicateEntity)
    assert ei.value.status_code == 1
