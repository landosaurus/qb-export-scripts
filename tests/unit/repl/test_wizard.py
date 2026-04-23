from __future__ import annotations

import logging
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest
from pytest_mock import MockerFixture

from qb_cli.io.format import Format
from qb_cli.ops.export_op import ExportResult
from qb_cli.ops.import_op import ImportResult
from qb_cli.ops.verify_op import VerifyResult
from qb_cli.repl.wizard import run_wizard
from qb_cli.utils.date_parse import year_range


def _canned(return_value: object) -> MagicMock:
    """Build a mock that matches the questionary API: foo(..).ask() -> value."""
    m = MagicMock()
    m.ask.return_value = return_value
    return m


def _canned_many(*values: object) -> MagicMock:
    """Build a mock whose .ask() returns a different value on each call, in order.

    Useful when the wizard asks multiple questions of the same type.
    """
    m = MagicMock()
    m.ask.side_effect = list(values)
    return m


def _fake_ctx() -> SimpleNamespace:
    return SimpleNamespace(logger=logging.getLogger("qb_cli.test"))


def test_exit_from_main_menu(mocker: MockerFixture) -> None:
    """User selects 'Exit' at main menu; wizard returns 0 without calling any op."""
    mocker.patch(
        "qb_cli.repl.wizard.questionary.select",
        return_value=_canned("Exit"),
    )
    exp = mocker.patch("qb_cli.repl.wizard.export")
    imp = mocker.patch("qb_cli.repl.wizard.import_")
    qry = mocker.patch("qb_cli.repl.wizard.query")
    ver = mocker.patch("qb_cli.repl.wizard.verify_entities")

    rc = run_wizard(_fake_ctx())
    assert rc == 0
    exp.assert_not_called()
    imp.assert_not_called()
    qry.assert_not_called()
    ver.assert_not_called()


def test_cancel_main_menu(mocker: MockerFixture) -> None:
    """questionary returns None (Ctrl-C) at main menu; wizard returns 0."""
    mocker.patch(
        "qb_cli.repl.wizard.questionary.select",
        return_value=_canned(None),
    )
    rc = run_wizard(_fake_ctx())
    assert rc == 0


def test_export_by_refs_happy_path(mocker: MockerFixture) -> None:
    """Export flow: main -> Export, entity -> invoice, mode -> refs, format -> CSV."""
    select_mock = mocker.patch(
        "qb_cli.repl.wizard.questionary.select",
        side_effect=[
            _canned("Export records"),  # main menu
            _canned("invoice"),  # entity
            _canned("Specific ref numbers (comma-separated)"),  # mode
            _canned("CSV"),  # format
            _canned("Exit"),  # back to main menu, then Exit
        ],
    )
    text_mock = mocker.patch(
        "qb_cli.repl.wizard.questionary.text",
        side_effect=[
            _canned("14396,14397"),  # refs
            _canned("out.csv"),  # output path
        ],
    )
    export_mock = mocker.patch(
        "qb_cli.repl.wizard.export",
        return_value=ExportResult(
            entity="invoice",
            count=2,
            output_path=Path("out.csv"),
            format=Format.CSV,
            pages=1,
        ),
    )

    rc = run_wizard(_fake_ctx())
    assert rc == 0
    assert export_mock.call_count == 1
    _, kwargs = export_mock.call_args
    assert kwargs["ref_numbers"] == ["14396", "14397"]
    assert kwargs["fmt"] is Format.CSV
    assert kwargs["output_path"] == Path("out.csv")
    # main menu invoked twice: once to start Export, once (Exit) after the op
    assert select_mock.call_count == 5
    assert text_mock.call_count == 2


def test_export_by_year(mocker: MockerFixture) -> None:
    """Export flow by year: dates derived from year_range(2025)."""
    mocker.patch(
        "qb_cli.repl.wizard.questionary.select",
        side_effect=[
            _canned("Export records"),
            _canned("invoice"),
            _canned("Year (all records from Jan 1 through Dec 31)"),
            _canned("JSON"),
            _canned("Exit"),
        ],
    )
    mocker.patch(
        "qb_cli.repl.wizard.questionary.text",
        side_effect=[
            _canned("2025"),  # year
            _canned("out.json"),  # output path
        ],
    )
    export_mock = mocker.patch(
        "qb_cli.repl.wizard.export",
        return_value=ExportResult(
            entity="invoice",
            count=0,
            output_path=Path("out.json"),
            format=Format.JSON,
        ),
    )

    rc = run_wizard(_fake_ctx())
    assert rc == 0
    expected_from, expected_to = year_range(2025)
    _, kwargs = export_mock.call_args
    assert kwargs["date_from"] == expected_from
    assert kwargs["date_to"] == expected_to
    assert kwargs["fmt"] is Format.JSON


def test_export_by_date_range(mocker: MockerFixture) -> None:
    """Export flow: user supplies explicit from/to dates."""
    mocker.patch(
        "qb_cli.repl.wizard.questionary.select",
        side_effect=[
            _canned("Export records"),
            _canned("invoice"),
            _canned("Date range"),
            _canned("CSV"),
            _canned("Exit"),
        ],
    )
    mocker.patch(
        "qb_cli.repl.wizard.questionary.text",
        side_effect=[
            _canned("2025-01-01"),  # date_from
            _canned("2025-12-31"),  # date_to
            _canned("out.csv"),  # output path
        ],
    )
    export_mock = mocker.patch(
        "qb_cli.repl.wizard.export",
        return_value=ExportResult(
            entity="invoice",
            count=0,
            output_path=Path("out.csv"),
            format=Format.CSV,
        ),
    )

    rc = run_wizard(_fake_ctx())
    assert rc == 0
    _, kwargs = export_mock.call_args
    assert kwargs["date_from"] == date(2025, 1, 1)
    assert kwargs["date_to"] == date(2025, 12, 31)


def test_import_dry_run(mocker: MockerFixture) -> None:
    """Import flow: JSON path, auto-detect format, dry run=True, on_duplicate=error, proceed=True."""
    mocker.patch(
        "qb_cli.repl.wizard.questionary.select",
        side_effect=[
            _canned("Import records"),
            _canned("invoice"),
            _canned("error"),  # on_duplicate
            _canned("Exit"),
        ],
    )
    mocker.patch(
        "qb_cli.repl.wizard.questionary.path",
        return_value=_canned("x.json"),
    )
    mocker.patch(
        "qb_cli.repl.wizard.questionary.confirm",
        side_effect=[
            _canned(True),  # dry run?
            _canned(True),  # proceed?
        ],
    )
    import_mock = mocker.patch(
        "qb_cli.repl.wizard.import_",
        return_value=ImportResult(entity="invoice", attempted=0, written=0),
    )

    rc = run_wizard(_fake_ctx())
    assert rc == 0
    assert import_mock.call_count == 1
    _, kwargs = import_mock.call_args
    assert kwargs["dry_run"] is True
    assert kwargs["on_duplicate"] == "error"
    assert kwargs["input_path"] == Path("x.json")
    assert kwargs["fmt"] is Format.JSON


def test_query_happy_path(
    mocker: MockerFixture, capsys: pytest.CaptureFixture[str]
) -> None:
    """Query flow: one ref, returns 1 entity; stdout contains the ref."""
    mocker.patch(
        "qb_cli.repl.wizard.questionary.select",
        side_effect=[
            _canned("Query records"),
            _canned("invoice"),
            _canned("Exit"),
        ],
    )
    mocker.patch(
        "qb_cli.repl.wizard.questionary.text",
        return_value=_canned("14396"),
    )
    fake_result = SimpleNamespace(
        ref_number="14396",
        txn_date=date(2025, 1, 15),
        total_amount="100.00",
    )
    query_mock = mocker.patch(
        "qb_cli.repl.wizard.query",
        return_value=[fake_result],
    )

    rc = run_wizard(_fake_ctx())
    assert rc == 0
    assert query_mock.call_count == 1
    _, kwargs = query_mock.call_args
    assert kwargs["ref_numbers"] == ["14396"]
    out = capsys.readouterr().out
    assert "14396" in out


def test_verify_happy_path(mocker: MockerFixture) -> None:
    """Verify flow: checkbox selects customers + items; each bucket gets a ref list."""
    mocker.patch(
        "qb_cli.repl.wizard.questionary.select",
        side_effect=[
            _canned("Verify entities exist"),
            _canned("Exit"),
        ],
    )
    mocker.patch(
        "qb_cli.repl.wizard.questionary.checkbox",
        return_value=_canned(["customers", "items"]),
    )
    mocker.patch(
        "qb_cli.repl.wizard.questionary.text",
        side_effect=[
            _canned("ACME"),  # customers
            _canned("40-RAG12"),  # items
        ],
    )
    verify_mock = mocker.patch(
        "qb_cli.repl.wizard.verify_entities",
        return_value=VerifyResult(
            found={
                "customer": {"ACME"},
                "vendor": set(),
                "item": {"40-RAG12"},
                "terms": set(),
            },
            missing={
                "customer": set(),
                "vendor": set(),
                "item": set(),
                "terms": set(),
            },
        ),
    )

    rc = run_wizard(_fake_ctx())
    assert rc == 0
    assert verify_mock.call_count == 1
    _, kwargs = verify_mock.call_args
    assert kwargs["customers"] == ["ACME"]
    assert kwargs["items"] == ["40-RAG12"]
    assert kwargs["vendors"] == ()
    assert kwargs["terms"] == ()


def test_op_exception_does_not_crash_wizard(mocker: MockerFixture) -> None:
    """If an op raises, wizard catches it, prints an error, loops back, user exits."""
    mocker.patch(
        "qb_cli.repl.wizard.questionary.select",
        side_effect=[
            _canned("Export records"),
            _canned("invoice"),
            _canned("Specific ref numbers (comma-separated)"),
            _canned("CSV"),
            _canned("Exit"),  # back to main menu after exception
        ],
    )
    mocker.patch(
        "qb_cli.repl.wizard.questionary.text",
        side_effect=[
            _canned("14396"),
            _canned("out.csv"),
        ],
    )
    mocker.patch(
        "qb_cli.repl.wizard.export",
        side_effect=RuntimeError("boom"),
    )

    rc = run_wizard(_fake_ctx())
    assert rc == 0
