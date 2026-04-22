from __future__ import annotations

from prompt_toolkit.document import Document

from qb_cli.repl.completers import QBCompleter


def test_verb_completion_from_empty() -> None:
    doc = Document("")
    completions = list(QBCompleter().get_completions(doc, None))
    labels = [c.text for c in completions]
    assert "export" in labels
    assert "import" in labels


def test_verb_completion_with_prefix() -> None:
    doc = Document("ex")
    completions = list(QBCompleter().get_completions(doc, None))
    labels = [c.text for c in completions]
    # ordering not asserted — just the set
    assert set(labels) == {"export", "exit"}


def test_entity_completion_after_export() -> None:
    doc = Document("export ")
    completions = list(QBCompleter().get_completions(doc, None))
    labels = [c.text for c in completions]
    assert "invoice" in labels
    assert "sales_order" in labels


def test_entity_prefix_after_export() -> None:
    doc = Document("export inv")
    completions = list(QBCompleter().get_completions(doc, None))
    labels = [c.text for c in completions]
    assert labels == ["invoice"]


def test_no_completion_for_unknown_verb() -> None:
    doc = Document("frobnicate ")
    assert list(QBCompleter().get_completions(doc, None)) == []


def test_no_completion_for_verify_second_token() -> None:
    doc = Document("verify ")
    # verify takes flags, not an entity positional — don't suggest entities
    completions = list(QBCompleter().get_completions(doc, None))
    assert [c.text for c in completions] == []
