from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from qb_cli.context import Context


_ENV_VARS = (
    "QB_CLI_COMPANY_FILE",
    "QB_CLI_DEFAULT_OUTPUT_DIR",
    "QB_CLI_DEFAULT_FORMAT",
    "QB_CLI_LOG_LEVEL",
)


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch):
    for var in _ENV_VARS:
        monkeypatch.delenv(var, raising=False)
    yield


def test_from_env_builds_context_with_logger_and_config(tmp_path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("QB_CLI_DEFAULT_OUTPUT_DIR", str(tmp_path))
    ctx = Context.from_env(company_file="X.QBW")
    assert ctx.config.company_file == "X.QBW"
    assert ctx.output_dir == tmp_path
    assert ctx.logger.name == "qb_cli"


def test_from_env_factory_opens_connection_with_company_file(monkeypatch: pytest.MonkeyPatch):
    fake_rp = MagicMock()
    fake_rp.BeginSession.return_value = "ticket-123"
    monkeypatch.setattr(
        "qb_cli.transport.connection._dispatch_request_processor",
        lambda: fake_rp,
    )
    ctx = Context.from_env(company_file="X.QBW")
    with ctx.connection_factory() as conn:
        assert conn.is_open
    fake_rp.BeginSession.assert_called_once()
    args, _ = fake_rp.BeginSession.call_args
    assert args[0] == "X.QBW"


def test_from_env_does_not_eagerly_open_connection(monkeypatch: pytest.MonkeyPatch):
    called = {"count": 0}

    def fake_dispatch():
        called["count"] += 1
        return MagicMock()

    monkeypatch.setattr(
        "qb_cli.transport.connection._dispatch_request_processor",
        fake_dispatch,
    )
    Context.from_env(company_file="Y.QBW")
    assert called["count"] == 0
