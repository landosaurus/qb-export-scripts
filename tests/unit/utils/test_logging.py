from __future__ import annotations

import io
import json
import logging

import pytest

from qb_cli.utils.logging import setup_logging


@pytest.fixture(autouse=True)
def _reset_logger():
    logger = logging.getLogger("qb_cli")
    original_handlers = list(logger.handlers)
    original_level = logger.level
    original_propagate = logger.propagate
    yield
    for h in list(logger.handlers):
        logger.removeHandler(h)
    for h in original_handlers:
        logger.addHandler(h)
    logger.setLevel(original_level)
    logger.propagate = original_propagate


def test_setup_logging_sets_debug_level():
    logger = setup_logging("DEBUG")
    assert logger.level == logging.DEBUG
    assert logger.level == 10


def test_setup_logging_is_idempotent():
    logger1 = setup_logging("INFO")
    handler_count = len(logger1.handlers)
    logger2 = setup_logging("INFO")
    assert logger1 is logger2
    assert len(logger2.handlers) == handler_count


def test_setup_logging_accepts_numeric_level():
    logger = setup_logging(logging.WARNING)
    assert logger.level == logging.WARNING


def test_setup_logging_json_output_contains_required_keys():
    from qb_cli.utils.logging import _JsonFormatter

    logger = setup_logging("DEBUG", json=True)
    buf = io.StringIO()
    for h in list(logger.handlers):
        logger.removeHandler(h)
    handler = logging.StreamHandler(buf)
    handler.setFormatter(_JsonFormatter())
    logger.addHandler(handler)

    logger.info("hello world")

    line = buf.getvalue().strip().splitlines()[-1]
    record = json.loads(line)
    assert record["message"] == "hello world"
    assert record["level"] == "INFO"
    assert "time" in record
    assert record["name"] == "qb_cli"
