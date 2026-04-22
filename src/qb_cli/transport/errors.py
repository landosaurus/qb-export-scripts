from dataclasses import dataclass


class QBError(Exception):
    """Base class for all QuickBooks-related errors raised by qb_cli."""


class QBConnectionError(QBError):
    """Failed to connect to or communicate with QuickBooks."""


class QBXMLParseError(QBError):
    """QBXML response could not be parsed."""


@dataclass
class QBStatusError(QBError):
    status_code: int
    status_message: str
    status_severity: str
    request_id: str

    def __str__(self) -> str:
        return f"QB status {self.status_code} ({self.status_severity}): {self.status_message} [request {self.request_id}]"


class QBEntityNotFound(QBStatusError):
    """Status 500 / 3120 - referenced entity does not exist."""


class QBDuplicateEntity(QBStatusError):
    """Status 3100 / 3270 - entity already exists (duplicate RefNumber or name)."""
