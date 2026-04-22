from __future__ import annotations

from pathlib import Path

import pytest

from qb_cli.config import Config, load_config


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


def test_missing_file_returns_defaults(tmp_path: Path):
    cfg = load_config(tmp_path / "nonexistent.toml")
    assert cfg == Config()


def test_loads_company_file_from_toml(tmp_path: Path):
    toml_path = tmp_path / "config.toml"
    toml_path.write_text('company_file = "C:/test.QBW"\n')
    cfg = load_config(toml_path)
    assert cfg.company_file == "C:/test.QBW"


def test_env_var_overrides_toml(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    toml_path = tmp_path / "config.toml"
    toml_path.write_text('company_file = "C:/from_toml.QBW"\n')
    monkeypatch.setenv("QB_CLI_COMPANY_FILE", "C:/from_env.QBW")
    cfg = load_config(toml_path)
    assert cfg.company_file == "C:/from_env.QBW"


def test_env_var_overrides_all_fields(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("QB_CLI_COMPANY_FILE", "X.QBW")
    monkeypatch.setenv("QB_CLI_DEFAULT_OUTPUT_DIR", str(tmp_path))
    monkeypatch.setenv("QB_CLI_DEFAULT_FORMAT", "json")
    monkeypatch.setenv("QB_CLI_LOG_LEVEL", "DEBUG")
    cfg = load_config(tmp_path / "nonexistent.toml")
    assert cfg.company_file == "X.QBW"
    assert cfg.default_output_dir == tmp_path
    assert cfg.default_format == "json"
    assert cfg.log_level == "DEBUG"
