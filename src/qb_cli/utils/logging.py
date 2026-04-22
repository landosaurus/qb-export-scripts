from __future__ import annotations

import json
import logging
import sys
from typing import Union


_LOGGER_NAME = "qb_cli"
_TEXT_FORMAT = "%(asctime)s %(levelname)-5s %(name)s: %(message)s"


class _JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        payload: dict[str, str] = {
            "time": self.formatTime(record),
            "level": record.levelname,
            "name": record.name,
            "message": record.getMessage(),
        }
        return json.dumps(payload)


def _coerce_level(level: Union[str, int]) -> int:
    if isinstance(level, int):
        return level
    resolved = logging.getLevelName(level.upper())
    if isinstance(resolved, int):
        return resolved
    raise ValueError(f"unknown log level: {level!r}")


def setup_logging(level: Union[str, int] = "INFO", json: bool = False) -> logging.Logger:
    """Configure the root qb_cli logger; return it.

    When json=True, emit single-line JSON records to stderr.
    Otherwise emit short human-readable text.
    """
    logger = logging.getLogger(_LOGGER_NAME)
    logger.setLevel(_coerce_level(level))

    for handler in list(logger.handlers):
        logger.removeHandler(handler)

    handler = logging.StreamHandler(sys.stderr)
    if json:
        handler.setFormatter(_JsonFormatter())
    else:
        handler.setFormatter(logging.Formatter(_TEXT_FORMAT))
    logger.addHandler(handler)
    logger.propagate = False
    return logger
