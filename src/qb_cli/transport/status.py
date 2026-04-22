from __future__ import annotations

from lxml import etree

from qb_cli.transport.errors import (
    QBDuplicateEntity,
    QBEntityNotFound,
    QBStatusError,
    QBXMLParseError,
)

_NOT_FOUND_CODES = {500, 3120}
_DUPLICATE_CODES = {3100, 3270}


def check_response_status(response_xml: str) -> None:
    """Scan a QBXML response for *Rs elements and raise on the first non-zero status.

    Warnings (statusCode != 0 with severity Warn) are also raised - callers that
    want to tolerate warnings should catch QBStatusError.
    """
    try:
        root = etree.fromstring(response_xml.encode("utf-8"))
    except etree.XMLSyntaxError as e:
        raise QBXMLParseError(f"invalid QBXML response: {e}") from e

    for rs in root.xpath("//*[local-name()='QBXMLMsgsRs']/*"):
        code = int(rs.get("statusCode", "0"))
        if code == 0:
            continue
        msg = rs.get("statusMessage", "")
        severity = rs.get("statusSeverity", "")
        req_id = rs.get("requestID", "")
        if code in _NOT_FOUND_CODES:
            raise QBEntityNotFound(code, msg, severity, req_id)
        if code in _DUPLICATE_CODES:
            raise QBDuplicateEntity(code, msg, severity, req_id)
        raise QBStatusError(code, msg, severity, req_id)
