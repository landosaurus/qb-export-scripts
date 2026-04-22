from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, ContextManager

from qb_cli.config import Config, load_config
from qb_cli.transport.connection import QBConnection
from qb_cli.utils.logging import setup_logging


ConnectionFactory = Callable[[], ContextManager[QBConnection]]


@dataclass
class Context:
    connection_factory: ConnectionFactory
    config: Config
    logger: logging.Logger
    output_dir: Path
    default_format: str
    json_output: bool
    dry_run: bool = False

    @classmethod
    def from_env(
        cls,
        *,
        company_file: str | None = None,
        config_path: str | None = None,
        log_level: str | None = None,
        json_output: bool = False,
    ) -> "Context":
        cfg = load_config(Path(config_path) if config_path is not None else None)
        if company_file is not None:
            cfg.company_file = company_file
        if log_level is not None:
            cfg.log_level = log_level
        logger = setup_logging(cfg.log_level, json=json_output)

        effective_company_file = cfg.company_file

        def factory() -> ContextManager[QBConnection]:
            return QBConnection(company_file=effective_company_file)

        return cls(
            connection_factory=factory,
            config=cfg,
            logger=logger,
            output_dir=cfg.default_output_dir,
            default_format=cfg.default_format,
            json_output=json_output,
        )
