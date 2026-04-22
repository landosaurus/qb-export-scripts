from __future__ import annotations

from typing import Optional

from qb_cli.transport.errors import QBConnectionError

_QB_FILE_MODE = 2  # qbFileOpenDoNotCare
_APP_ID = ""
_APP_NAME = "qb-cli"
_CONN_TYPE = 1  # localQBD


def _dispatch_request_processor() -> object:
    """Late import so non-Windows CI can still import this module for unit tests."""
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as e:
        raise QBConnectionError(
            "pywin32 is not installed. Install qb-cli with the 'windows' extra "
            "on the machine where QuickBooks is running."
        ) from e
    return win32com.client.Dispatch("QBXMLRP2.RequestProcessor")


class QBConnection:
    """Context manager around QBXMLRP2.RequestProcessor.

    Usage:
        with QBConnection(company_file="C:/path/to/file.QBW") as conn:
            response_xml = conn.send(request_xml)
    """

    def __init__(self, company_file: str = "") -> None:
        self._company_file = company_file
        self._rp: Optional[object] = None
        self._ticket: Optional[str] = None

    @property
    def is_open(self) -> bool:
        return self._ticket is not None

    def __enter__(self) -> "QBConnection":
        self._rp = _dispatch_request_processor()
        try:
            self._rp.OpenConnection2(_APP_ID, _APP_NAME, _CONN_TYPE)
            self._ticket = self._rp.BeginSession(self._company_file, _QB_FILE_MODE)
        except Exception as e:
            raise QBConnectionError(f"failed to open QuickBooks session: {e}") from e
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._ticket is not None and self._rp is not None:
            try:
                self._rp.EndSession(self._ticket)
            finally:
                self._ticket = None
                try:
                    self._rp.CloseConnection()
                except Exception:
                    pass
                self._rp = None

    def send(self, qbxml_request: str) -> str:
        if self._ticket is None or self._rp is None:
            raise QBConnectionError("QBConnection is not open")
        try:
            return self._rp.ProcessRequest(self._ticket, qbxml_request)
        except Exception as e:
            raise QBConnectionError(f"ProcessRequest failed: {e}") from e
