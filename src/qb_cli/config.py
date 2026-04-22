from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import cast

try:
    import tomllib
except ImportError:  # pragma: no cover
    import tomli as tomllib  # type: ignore[no-redef]


@dataclass
class Config:
    company_file: str = ""
    default_output_dir: Path = field(default_factory=Path.cwd)
    default_format: str = "csv"
    log_level: str = "INFO"


def _default_config_path() -> Path:
    """XDG on macOS/Linux, APPDATA on Windows."""
    if os.name == "nt":
        appdata = os.environ.get("APPDATA")
        if appdata:
            return Path(appdata) / "qb_cli" / "config.toml"
        return Path.home() / "AppData" / "Roaming" / "qb_cli" / "config.toml"
    xdg = os.environ.get("XDG_CONFIG_HOME")
    base = Path(xdg) if xdg else Path.home() / ".config"
    return base / "qb_cli" / "config.toml"


def _apply_env_overrides(cfg: Config) -> Config:
    env_company = os.environ.get("QB_CLI_COMPANY_FILE")
    if env_company is not None:
        cfg.company_file = env_company
    env_output = os.environ.get("QB_CLI_DEFAULT_OUTPUT_DIR")
    if env_output is not None:
        cfg.default_output_dir = Path(env_output)
    env_fmt = os.environ.get("QB_CLI_DEFAULT_FORMAT")
    if env_fmt is not None:
        cfg.default_format = env_fmt
    env_level = os.environ.get("QB_CLI_LOG_LEVEL")
    if env_level is not None:
        cfg.log_level = env_level
    return cfg


def load_config(path: Path | None = None) -> Config:
    """Load config from TOML at `path` or the default location. Missing file → defaults."""
    cfg = Config()
    target = path if path is not None else _default_config_path()
    if target.is_file():
        with target.open("rb") as fh:
            data = cast(dict[str, object], tomllib.load(fh))
        company_file = data.get("company_file")
        if isinstance(company_file, str):
            cfg.company_file = company_file
        default_output_dir = data.get("default_output_dir")
        if isinstance(default_output_dir, str):
            cfg.default_output_dir = Path(default_output_dir)
        default_format = data.get("default_format")
        if isinstance(default_format, str):
            cfg.default_format = default_format
        log_level = data.get("log_level")
        if isinstance(log_level, str):
            cfg.log_level = log_level
    return _apply_env_overrides(cfg)
