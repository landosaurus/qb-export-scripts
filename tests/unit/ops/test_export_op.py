from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from qb_cli.io.format import Format
from qb_cli.ops.export_op import ExportResult, export
from tests.unit.ops._fakes import make_factory


_FIXTURE_DIR = Path(__file__).resolve().parents[2] / "fixtures" / "qbxml_responses"


def _ctx_with(response_xml: str) -> SimpleNamespace:
    return SimpleNamespace(connection_factory=make_factory(response_xml))


def test_export_invoice_to_csv(tmp_path: Path):
    fixture = (_FIXTURE_DIR / "invoice_query_sample.xml").read_text()
    ctx = _ctx_with(fixture)
    out = tmp_path / "inv.csv"

    result = export(ctx, "invoice", output_path=out)

    assert isinstance(result, ExportResult)
    assert result.count == 2
    assert result.output_path == out
    assert result.format is Format.CSV
    assert out.exists()
    assert "row_type" in out.read_text().splitlines()[0]


def test_export_sales_order_to_json(tmp_path: Path):
    fixture = (_FIXTURE_DIR / "sales_order_query_sample.xml").read_text()
    ctx = _ctx_with(fixture)
    out = tmp_path / "so.json"

    result = export(ctx, "sales_order", output_path=out)

    assert result.count == 2
    assert result.format is Format.JSON
    data = json.loads(out.read_text())
    assert isinstance(data, list)
    assert len(data) == 2


def test_export_purchase_order_sends_query_with_filters(tmp_path: Path):
    fixture = (_FIXTURE_DIR / "purchase_order_query_sample.xml").read_text()
    factory = make_factory(fixture)
    ctx = SimpleNamespace(connection_factory=factory)
    out = tmp_path / "po.json"

    result = export(
        ctx,
        "purchase_order",
        ref_numbers=["7740", "7741"],
        output_path=out,
    )
    assert result.count == 2


def test_export_unknown_entity_raises(tmp_path: Path):
    ctx = _ctx_with("")
    with pytest.raises(KeyError):
        export(ctx, "widget", output_path=tmp_path / "out.csv")
