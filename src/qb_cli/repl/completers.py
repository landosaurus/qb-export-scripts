from __future__ import annotations

from prompt_toolkit.completion import CompleteEvent, Completer, Completion
from prompt_toolkit.document import Document


VERBS: list[str] = [
    "export",
    "import",
    "query",
    "verify",
    "help",
    "set",
    "unset",
    "status",
    "connect",
    "disconnect",
    "exit",
    "quit",
]
ENTITIES: list[str] = ["invoice", "sales_order", "purchase_order"]


class QBCompleter(Completer):
    """Context-aware completer for the REPL.

    Level 0 (no tokens yet): complete VERBS.
    Level 1 after a verb that takes an entity (export/import/query): complete ENTITIES.
    Otherwise: no completions.
    """

    ENTITY_TAKING_VERBS: set[str] = {"export", "import", "query"}

    def get_completions(
        self, document: Document, complete_event: CompleteEvent | None
    ) -> list[Completion]:
        text = document.text_before_cursor
        tokens = text.split()
        # Determine if the cursor is at the start of a new token (text ends with whitespace or is empty).
        at_word_boundary = text == "" or text.endswith((" ", "\t"))
        position = len(tokens) if at_word_boundary else len(tokens) - 1
        current = "" if at_word_boundary else tokens[-1]

        if position == 0:
            return self._match(current, VERBS)
        if position == 1 and tokens and tokens[0] in self.ENTITY_TAKING_VERBS:
            return self._match(current, ENTITIES)
        return []

    @staticmethod
    def _match(current: str, candidates: list[str]) -> list[Completion]:
        return [
            Completion(c, start_position=-len(current))
            for c in candidates
            if c.startswith(current)
        ]
