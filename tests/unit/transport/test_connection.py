import pytest
from qb_cli.transport.connection import QBConnection
from qb_cli.transport.errors import QBConnectionError


class FakeRequestProcessor:
    def __init__(self):
        self.opened = False
        self.session_ticket = None
        self.sent = []

    def OpenConnection2(self, app_id, app_name, conn_type):
        self.opened = True
        self.app_name = app_name

    def BeginSession(self, company_file, mode):
        assert self.opened
        self.session_ticket = "TICKET-1"
        self.company_file = company_file
        self.mode = mode
        return self.session_ticket

    def ProcessRequest(self, ticket, request):
        assert ticket == self.session_ticket
        self.sent.append(request)
        return "<QBXML><response/></QBXML>"

    def EndSession(self, ticket):
        assert ticket == self.session_ticket
        self.session_ticket = None

    def CloseConnection(self):
        self.opened = False


def test_context_manager_opens_and_closes(mocker):
    fake = FakeRequestProcessor()
    mocker.patch("qb_cli.transport.connection._dispatch_request_processor", return_value=fake)

    with QBConnection(company_file="") as conn:
        assert conn.is_open
        assert fake.opened
        assert fake.session_ticket == "TICKET-1"

    assert not fake.opened
    assert fake.session_ticket is None


def test_send_round_trip(mocker):
    fake = FakeRequestProcessor()
    mocker.patch("qb_cli.transport.connection._dispatch_request_processor", return_value=fake)

    with QBConnection() as conn:
        response = conn.send("<QBXML><request/></QBXML>")

    assert response == "<QBXML><response/></QBXML>"
    assert fake.sent == ["<QBXML><request/></QBXML>"]


def test_open_failure_raises_connection_error(mocker):
    fake = FakeRequestProcessor()
    def boom(*a, **kw):
        raise OSError("COM error")
    fake.OpenConnection2 = boom
    mocker.patch("qb_cli.transport.connection._dispatch_request_processor", return_value=fake)

    with pytest.raises(QBConnectionError):
        with QBConnection():
            pass
