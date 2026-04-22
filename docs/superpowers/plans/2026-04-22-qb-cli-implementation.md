# qb-cli Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a unified `qb` tool (REPL + non-interactive CLI) for QuickBooks Desktop supporting export + import for 15 entity types with CSV/JSON round-trip, pre-flight verification, and duplicate detection.

**Architecture:** Layered — `transport` (COM+QBXML wire) → `qbxml` (per-entity builders/parsers) → `models` (pydantic) → `io` (CSV/JSON serializers) + `ops` (export/import/query/verify/dedupe) → `cli` (click) → `repl` (prompt_toolkit). Every arrow points one way; lower layers never import higher layers.

**Tech Stack:** Python 3.10+, click, prompt_toolkit, pydantic v2, lxml, pywin32 (Windows runtime), pytest, hypothesis.

**Spec reference:** `docs/superpowers/specs/2026-04-22-qb-cli-repl-design.md`

**qb-mcp reference root:** `/Users/lang/Documents/coding_temp/qb-mcp/src/qb_mcp/` — used for **reading only**. Do not import, vendor, or copy code from qb-mcp. Reference it when you need to remember the exact shape of a QBXML request for a given entity. The file `qb-mcp/src/qb_mcp/tools/<entity>s/<entity>_query.py` shows how qb-mcp builds queries; `_add.py` and `_mod.py` show the mutation shapes.

---

## Conventions for every task

- **TDD.** Write the failing test first. Run it. Implement the minimal code to pass. Run again. Commit.
- **Commit message shape:** `<area>: <subject>` (e.g. `transport: add QBConnection context manager`, `qbxml(invoice): build query for ref number`).
- **Commit cadence:** one commit per task, after tests pass. Never commit with failing tests. Never `--no-verify`.
- **Type hints everywhere.** Never use `typing.Any` or cast to `Any`; if you cannot figure out a type, use context7 to look up the authoritative type for the library in question. This is a hard project rule (per user's global `CLAUDE.md`).
- **No comments that restate what code does.** Only write a comment when the *why* is non-obvious.
- **Never assume QB is reachable.** Every test in this plan mocks the transport layer unless marked `pytest.mark.integration`.

## Parallelization map

Task groups marked with **‖** are parallelizable: dispatch one subagent per item.

- Tasks 1–3 are sequential (scaffolding, transport, model base).
- Tasks 4.x (per-entity models) are all **‖**.
- Tasks 5.x (per-entity QBXML builders/parsers) are all **‖** but depend on their matching 4.x completing.
- Tasks 6.x (IO serializers) are **‖** with 5.x.
- Tasks 7–9 are sequential on top (ops → cli → repl).

---

## Entity list (canonical order for all per-entity tasks)

| # | Entity key | qbxml root | line items? | subtypes? |
|---|---|---|---|---|
| 1 | `invoice` | `InvoiceQuery/Add/ModRq` | yes | no |
| 2 | `sales_order` | `SalesOrderQuery/Add/ModRq` | yes | no |
| 3 | `purchase_order` | `PurchaseOrderQuery/Add/ModRq` | yes | no |
| 4 | `bill` | `BillQuery/Add/ModRq` | yes (expense + item) | no |
| 5 | `estimate` | `EstimateQuery/Add/ModRq` | yes | no |
| 6 | `sales_receipt` | `SalesReceiptQuery/Add/ModRq` | yes | no |
| 7 | `credit_memo` | `CreditMemoQuery/Add/ModRq` | yes | no |
| 8 | `receive_payment` | `ReceivePaymentQuery/Add/ModRq` | yes (applied-to) | no |
| 9 | `check` | `CheckQuery/Add/ModRq` | yes | no |
| 10 | `deposit` | `DepositQuery/Add/ModRq` | yes | no |
| 11 | `customer` | `CustomerQuery/Add/ModRq` | no | no |
| 12 | `vendor` | `VendorQuery/Add/ModRq` | no | no |
| 13 | `item` | `ItemQueryRq` + `ItemInventoryAdd/ModRq`, `ItemServiceAdd/ModRq`, `ItemNonInventoryAdd/ModRq`, `ItemOtherChargeAdd/ModRq` | no | yes (4 subtypes) |
| 14 | `price_level` | `PriceLevelQuery/Add/ModRq` | no | no |
| 15 | `ship_to` | `CustomerQueryRq` (embedded under customer) + `CustomerModRq` for add/update | no | no |

---

## Task 1: Scaffold the project

**Files:**
- Create: `pyproject.toml` (new at repo root; replace dev usage of `requirements.txt`)
- Create: `src/qb_cli/__init__.py`
- Create: `src/qb_cli/version.py`
- Create: `src/qb_cli/__main__.py`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`
- Create: `.gitignore` additions (`.venv/`, `__pycache__/`, `.pytest_cache/`, `dist/`, `build/`, `*.egg-info/`)

- [ ] **Step 1.1: Create `pyproject.toml`**

```toml
[project]
name = "qb-cli"
version = "0.1.0"
description = "Interactive REPL + non-interactive CLI for QuickBooks Desktop"
requires-python = ">=3.10"
dependencies = [
    "click>=8.1",
    "prompt-toolkit>=3.0",
    "pydantic>=2.6",
    "lxml>=4.9",
    "tomli>=2.0;python_version<'3.11'",
]

[project.optional-dependencies]
windows = ["pywin32>=306"]
dev = [
    "pytest>=8.0",
    "pytest-mock>=3.12",
    "hypothesis>=6.98",
    "responses>=0.24",
]

[project.scripts]
qb = "qb_cli.__main__:main"

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["src/qb_cli"]

[tool.pytest.ini_options]
testpaths = ["tests"]
python_files = ["test_*.py"]
addopts = ["-v", "--tb=short", "--strict-markers", "-ra"]
markers = [
    "unit: unit tests (default)",
    "integration: requires live QuickBooks Desktop connection",
]
```

- [ ] **Step 1.2: Create `src/qb_cli/version.py`**

```python
__version__ = "0.1.0"
```

- [ ] **Step 1.3: Create `src/qb_cli/__init__.py`**

```python
from qb_cli.version import __version__

__all__ = ["__version__"]
```

- [ ] **Step 1.4: Create placeholder `src/qb_cli/__main__.py`**

```python
def main() -> None:
    raise SystemExit("qb_cli entry point not yet wired — see cli/root.py (task 8)")


if __name__ == "__main__":
    main()
```

- [ ] **Step 1.5: Create `tests/conftest.py`**

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
```

- [ ] **Step 1.6: Create `tests/__init__.py`** (empty file)

- [ ] **Step 1.7: Append to `.gitignore`**

```
.venv/
__pycache__/
*.pyc
.pytest_cache/
dist/
build/
*.egg-info/
.coverage
```

- [ ] **Step 1.8: Install the project and verify pytest discovery**

Run:
```bash
python -m pip install -e '.[dev]'
python -m pytest --collect-only
```
Expected: `collected 0 items` (no tests yet, no errors). If `lxml` or `pywin32` fails to install on macOS, that's fine — `pywin32` is `windows`-extra only; `lxml` should install everywhere.

- [ ] **Step 1.9: Commit**

```bash
git add pyproject.toml src/qb_cli tests .gitignore
git commit -m "scaffold: add qb_cli package skeleton and pyproject"
```

---

## Task 2: Archive existing scripts

**Files:**
- Create: `archive/README-archived.md`
- Move: `qb_inv.py`, `qb_so.py`, `qb_po.py`, `qb_bills.py`, `qb_price_level.py`, `qb_shipto.py`, `example_customer_query.py`, `example_purchase_order_add.py`, `test_ship_to_query.py`, `CUSTOMER_QUERY_README.md`, `ESTIMATE_QUERY_REFERENCE.md` → `archive/`
- Modify: root `README.md` (stub for now; full rewrite in task 10)

- [ ] **Step 2.1: Create `archive/` and move scripts with `git mv`**

```bash
mkdir -p archive
git mv qb_inv.py qb_so.py qb_po.py qb_bills.py qb_price_level.py qb_shipto.py archive/
git mv example_customer_query.py example_purchase_order_add.py test_ship_to_query.py archive/
git mv CUSTOMER_QUERY_README.md ESTIMATE_QUERY_REFERENCE.md archive/
```

- [ ] **Step 2.2: Create `archive/README-archived.md`**

```markdown
# Archived standalone scripts

These are the original one-off QuickBooks export scripts. They are preserved
here as a reference and for emergency use while the new `qb` CLI is under
development. The new tool lives in `src/qb_cli/` and is documented in the
root `README.md`.

Each script is standalone — it issues its own QBXML and writes a CSV.
They continue to work on the Windows QuickBooks machine as long as pywin32
is installed.

Do not add new features here. New work belongs in `src/qb_cli/`.
```

- [ ] **Step 2.3: Stub root `README.md`**

```markdown
# qb-cli

Interactive REPL + non-interactive CLI for QuickBooks Desktop.

**Status:** under active development. See `docs/superpowers/specs/2026-04-22-qb-cli-repl-design.md` for the design and `docs/superpowers/plans/2026-04-22-qb-cli-implementation.md` for the in-flight implementation plan.

The original standalone export scripts are preserved under `archive/`.
```

- [ ] **Step 2.4: Verify**

```bash
git status
ls archive/
```
Expected: `archive/` contains the moved scripts; working tree has no uncommitted non-archive changes.

- [ ] **Step 2.5: Commit**

```bash
git add -A archive/ README.md
git commit -m "archive: move standalone export scripts to archive/"
```

---

## Task 3: Transport layer

Builds the low-level QB connection and QBXML send/receive primitives. The only layer in the codebase that touches `win32com.client`. All tests below use `mocker` (pytest-mock) to fake the COM Dispatch.

### Task 3.1: Error hierarchy

**Files:**
- Create: `src/qb_cli/transport/__init__.py` (empty)
- Create: `src/qb_cli/transport/errors.py`
- Test: `tests/unit/transport/test_errors.py`

- [ ] **Step 3.1.1: Write failing tests**

```python
# tests/unit/transport/test_errors.py
import pytest
from qb_cli.transport.errors import (
    QBError,
    QBConnectionError,
    QBXMLParseError,
    QBStatusError,
    QBEntityNotFound,
    QBDuplicateEntity,
)


def test_qb_status_error_carries_fields():
    err = QBStatusError(
        status_code=3100,
        status_message="name already exists",
        status_severity="Error",
        request_id="1",
    )
    assert err.status_code == 3100
    assert "3100" in str(err)
    assert "name already exists" in str(err)


def test_entity_not_found_is_status_error():
    err = QBEntityNotFound(status_code=500, status_message="not found", status_severity="Error", request_id="1")
    assert isinstance(err, QBStatusError)
    assert isinstance(err, QBError)


def test_duplicate_entity_is_status_error():
    err = QBDuplicateEntity(status_code=3100, status_message="dup", status_severity="Error", request_id="1")
    assert isinstance(err, QBStatusError)


def test_base_error_hierarchy():
    assert issubclass(QBConnectionError, QBError)
    assert issubclass(QBXMLParseError, QBError)
    assert issubclass(QBStatusError, QBError)
```

- [ ] **Step 3.1.2: Run — expect fail**

`pytest tests/unit/transport/test_errors.py -v` → ImportError.

- [ ] **Step 3.1.3: Implement**

```python
# src/qb_cli/transport/errors.py
from dataclasses import dataclass


class QBError(Exception):
    """Base class for all QuickBooks-related errors raised by qb_cli."""


class QBConnectionError(QBError):
    """Failed to connect to or communicate with QuickBooks."""


class QBXMLParseError(QBError):
    """QBXML response could not be parsed."""


@dataclass
class QBStatusError(QBError):
    status_code: int
    status_message: str
    status_severity: str
    request_id: str

    def __str__(self) -> str:
        return f"QB status {self.status_code} ({self.status_severity}): {self.status_message} [request {self.request_id}]"


class QBEntityNotFound(QBStatusError):
    """Status 500 / 3120 — referenced entity does not exist."""


class QBDuplicateEntity(QBStatusError):
    """Status 3100 / 3270 — entity already exists (duplicate RefNumber or name)."""
```

- [ ] **Step 3.1.4: Run — expect pass**

`pytest tests/unit/transport/test_errors.py -v` → 4 passed.

- [ ] **Step 3.1.5: Commit**

```bash
git add src/qb_cli/transport tests/unit/transport/test_errors.py
git commit -m "transport: add QB error hierarchy"
```

### Task 3.2: QBConnection context manager

**Files:**
- Create: `src/qb_cli/transport/connection.py`
- Test: `tests/unit/transport/test_connection.py`

- [ ] **Step 3.2.1: Write failing tests**

```python
# tests/unit/transport/test_connection.py
import pytest
from unittest.mock import MagicMock, patch
from qb_cli.transport.connection import QBConnection
from qb_cli.transport.errors import QBConnectionError


class FakeRequestProcessor:
    def __init__(self):
        self.opened = False
        self.session_ticket = None
        self.sent = []

    def OpenConnection2(self, app_id, app_name, conn_type):
        self.opened = True
        self.app_name = app_name

    def BeginSession(self, company_file, mode):
        assert self.opened
        self.session_ticket = "TICKET-1"
        self.company_file = company_file
        self.mode = mode
        return self.session_ticket

    def ProcessRequest(self, ticket, request):
        assert ticket == self.session_ticket
        self.sent.append(request)
        return "<QBXML><response/></QBXML>"

    def EndSession(self, ticket):
        assert ticket == self.session_ticket
        self.session_ticket = None

    def CloseConnection(self):
        self.opened = False


def test_context_manager_opens_and_closes(mocker):
    fake = FakeRequestProcessor()
    mocker.patch("qb_cli.transport.connection._dispatch_request_processor", return_value=fake)

    with QBConnection(company_file="") as conn:
        assert conn.is_open
        assert fake.opened
        assert fake.session_ticket == "TICKET-1"

    assert not fake.opened
    assert fake.session_ticket is None


def test_send_round_trip(mocker):
    fake = FakeRequestProcessor()
    mocker.patch("qb_cli.transport.connection._dispatch_request_processor", return_value=fake)

    with QBConnection() as conn:
        response = conn.send("<QBXML><request/></QBXML>")

    assert response == "<QBXML><response/></QBXML>"
    assert fake.sent == ["<QBXML><request/></QBXML>"]


def test_open_failure_raises_connection_error(mocker):
    fake = FakeRequestProcessor()
    def boom(*a, **kw):
        raise OSError("COM error")
    fake.OpenConnection2 = boom
    mocker.patch("qb_cli.transport.connection._dispatch_request_processor", return_value=fake)

    with pytest.raises(QBConnectionError):
        with QBConnection():
            pass
```

- [ ] **Step 3.2.2: Run — expect fail**

- [ ] **Step 3.2.3: Implement `QBConnection`**

```python
# src/qb_cli/transport/connection.py
from __future__ import annotations

from typing import Optional

from qb_cli.transport.errors import QBConnectionError

# Mode 2 = qbFileOpenDoNotCare
_QB_FILE_MODE = 2
_APP_ID = ""
_APP_NAME = "qb-cli"
# Connection type 1 = localQBD (Desktop)
_CONN_TYPE = 1


def _dispatch_request_processor():
    """Late import so non-Windows CI can still import this module for unit tests."""
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as e:
        raise QBConnectionError(
            "pywin32 is not installed. Install qb-cli with the 'windows' extra "
            "on the machine where QuickBooks is running."
        ) from e
    return win32com.client.Dispatch("QBXMLRP2.RequestProcessor")


class QBConnection:
    """Context manager around QBXMLRP2.RequestProcessor.

    Usage:
        with QBConnection(company_file="C:/path/to/file.QBW") as conn:
            response_xml = conn.send(request_xml)
    """

    def __init__(self, company_file: str = "") -> None:
        self._company_file = company_file
        self._rp = None
        self._ticket: Optional[str] = None

    @property
    def is_open(self) -> bool:
        return self._ticket is not None

    def __enter__(self) -> "QBConnection":
        self._rp = _dispatch_request_processor()
        try:
            self._rp.OpenConnection2(_APP_ID, _APP_NAME, _CONN_TYPE)
            self._ticket = self._rp.BeginSession(self._company_file, _QB_FILE_MODE)
        except Exception as e:
            raise QBConnectionError(f"failed to open QuickBooks session: {e}") from e
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._ticket is not None and self._rp is not None:
            try:
                self._rp.EndSession(self._ticket)
            finally:
                self._ticket = None
                try:
                    self._rp.CloseConnection()
                except Exception:
                    pass
                self._rp = None

    def send(self, qbxml_request: str) -> str:
        if self._ticket is None or self._rp is None:
            raise QBConnectionError("QBConnection is not open")
        try:
            return self._rp.ProcessRequest(self._ticket, qbxml_request)
        except Exception as e:
            raise QBConnectionError(f"ProcessRequest failed: {e}") from e
```

- [ ] **Step 3.2.4: Run — expect pass**

- [ ] **Step 3.2.5: Commit**

```bash
git add src/qb_cli/transport/connection.py tests/unit/transport/test_connection.py
git commit -m "transport: add QBConnection context manager with mocked COM tests"
```

### Task 3.3: Response status extraction

**Files:**
- Create: `src/qb_cli/transport/status.py`
- Test: `tests/unit/transport/test_status.py`

- [ ] **Step 3.3.1: Write failing tests**

```python
# tests/unit/transport/test_status.py
import pytest
from qb_cli.transport.status import check_response_status
from qb_cli.transport.errors import QBEntityNotFound, QBDuplicateEntity, QBStatusError


OK_XML = """<?xml version="1.0"?>
<QBXML><QBXMLMsgsRs>
  <InvoiceQueryRs requestID="1" statusCode="0" statusSeverity="Info" statusMessage="Status OK"/>
</QBXMLMsgsRs></QBXML>"""

NOT_FOUND_XML = """<?xml version="1.0"?>
<QBXML><QBXMLMsgsRs>
  <InvoiceQueryRs requestID="1" statusCode="500" statusSeverity="Error" statusMessage="No matching"/>
</QBXMLMsgsRs></QBXML>"""

DUP_XML = """<?xml version="1.0"?>
<QBXML><QBXMLMsgsRs>
  <InvoiceAddRs requestID="1" statusCode="3100" statusSeverity="Error" statusMessage="name already exists"/>
</QBXMLMsgsRs></QBXML>"""

OTHER_ERR_XML = """<?xml version="1.0"?>
<QBXML><QBXMLMsgsRs>
  <InvoiceAddRs requestID="1" statusCode="1" statusSeverity="Warn" statusMessage="something"/>
</QBXMLMsgsRs></QBXML>"""


def test_ok_returns_normally():
    check_response_status(OK_XML)


def test_not_found_raises_entity_not_found():
    with pytest.raises(QBEntityNotFound) as ei:
        check_response_status(NOT_FOUND_XML)
    assert ei.value.status_code == 500


def test_duplicate_raises_duplicate_entity():
    with pytest.raises(QBDuplicateEntity) as ei:
        check_response_status(DUP_XML)
    assert ei.value.status_code == 3100


def test_other_error_raises_plain_status_error():
    with pytest.raises(QBStatusError) as ei:
        check_response_status(OTHER_ERR_XML)
    assert not isinstance(ei.value, QBDuplicateEntity)
    assert ei.value.status_code == 1
```

- [ ] **Step 3.3.2: Run — expect fail**

- [ ] **Step 3.3.3: Implement**

```python
# src/qb_cli/transport/status.py
from __future__ import annotations

from lxml import etree

from qb_cli.transport.errors import (
    QBDuplicateEntity,
    QBEntityNotFound,
    QBStatusError,
    QBXMLParseError,
)

_NOT_FOUND_CODES = {500, 3120}
_DUPLICATE_CODES = {3100, 3270}


def check_response_status(response_xml: str) -> None:
    """Scan a QBXML response for *Rs elements and raise on the first non-zero status.

    Warnings (statusCode != 0 with severity Warn) are also raised — callers that
    want to tolerate warnings should catch QBStatusError.
    """
    try:
        root = etree.fromstring(response_xml.encode("utf-8"))
    except etree.XMLSyntaxError as e:
        raise QBXMLParseError(f"invalid QBXML response: {e}") from e

    for rs in root.xpath("//*[local-name()='QBXMLMsgsRs']/*"):
        code = int(rs.get("statusCode", "0"))
        if code == 0:
            continue
        msg = rs.get("statusMessage", "")
        severity = rs.get("statusSeverity", "")
        req_id = rs.get("requestID", "")
        if code in _NOT_FOUND_CODES:
            raise QBEntityNotFound(code, msg, severity, req_id)
        if code in _DUPLICATE_CODES:
            raise QBDuplicateEntity(code, msg, severity, req_id)
        raise QBStatusError(code, msg, severity, req_id)
```

- [ ] **Step 3.3.4: Run — expect pass**

- [ ] **Step 3.3.5: Commit**

```bash
git add src/qb_cli/transport/status.py tests/unit/transport/test_status.py
git commit -m "transport: add response status extraction"
```

---

## Task 4: Model foundation (base + shared submodels)

**Files:**
- Create: `src/qb_cli/models/__init__.py`
- Create: `src/qb_cli/models/base.py`
- Create: `src/qb_cli/models/shared.py`
- Test: `tests/unit/models/test_base.py`
- Test: `tests/unit/models/test_shared.py`

- [ ] **Step 4.1: Write failing tests**

```python
# tests/unit/models/test_base.py
import pytest
from pydantic import ValidationError
from qb_cli.models.base import BaseEntity


class Demo(BaseEntity):
    ref_number: str | None = None
    memo: str | None = None


def test_populate_by_name_and_alias():
    # Accepts both Python and QB casing on input
    e1 = Demo(ref_number="1", memo="m")
    e2 = Demo.model_validate({"RefNumber": "2", "Memo": "n"})
    assert e1.ref_number == "1"
    assert e2.ref_number == "2"


def test_extra_fields_forbidden():
    with pytest.raises(ValidationError):
        Demo.model_validate({"RefNumber": "1", "NotARealField": "x"})


def test_dumps_with_qb_aliases_by_default():
    e = Demo(ref_number="1", memo="m")
    payload = e.model_dump(by_alias=True, exclude_none=True)
    assert payload == {"RefNumber": "1", "Memo": "m"}
```

```python
# tests/unit/models/test_shared.py
import pytest
from decimal import Decimal
from qb_cli.models.shared import Address, Ref, quantize_money


def test_address_render_multiline():
    a = Address(addr1="100 Main", city="Seattle", state="WA", postal_code="98101")
    assert a.render_multiline() == "100 Main\nSeattle, WA 98101"


def test_address_accepts_qb_casing():
    a = Address.model_validate({"Addr1": "100 Main", "City": "X", "State": "Y", "PostalCode": "1"})
    assert a.addr1 == "100 Main"


def test_ref_requires_full_name_or_list_id():
    r = Ref(full_name="ACME Corp")
    assert r.full_name == "ACME Corp"


def test_money_quantization():
    assert quantize_money(Decimal("10.1")) == Decimal("10.10")
    assert quantize_money(Decimal("10.123")) == Decimal("10.12")
```

- [ ] **Step 4.2: Run — expect fail**

- [ ] **Step 4.3: Implement `BaseEntity`**

```python
# src/qb_cli/models/__init__.py
from qb_cli.models.base import BaseEntity

__all__ = ["BaseEntity"]
```

```python
# src/qb_cli/models/base.py
from __future__ import annotations

from pydantic import BaseModel, ConfigDict


class BaseEntity(BaseModel):
    model_config = ConfigDict(
        populate_by_name=True,
        extra="forbid",
        str_strip_whitespace=True,
    )
```

- [ ] **Step 4.4: Implement `shared.py`**

```python
# src/qb_cli/models/shared.py
from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from typing import Optional

from pydantic import Field

from qb_cli.models.base import BaseEntity


def quantize_money(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


class Address(BaseEntity):
    addr1: Optional[str] = Field(default=None, alias="Addr1")
    addr2: Optional[str] = Field(default=None, alias="Addr2")
    addr3: Optional[str] = Field(default=None, alias="Addr3")
    addr4: Optional[str] = Field(default=None, alias="Addr4")
    addr5: Optional[str] = Field(default=None, alias="Addr5")
    city: Optional[str] = Field(default=None, alias="City")
    state: Optional[str] = Field(default=None, alias="State")
    postal_code: Optional[str] = Field(default=None, alias="PostalCode")
    country: Optional[str] = Field(default=None, alias="Country")
    note: Optional[str] = Field(default=None, alias="Note")

    def render_multiline(self) -> str:
        lines: list[str] = []
        for line in (self.addr1, self.addr2, self.addr3, self.addr4, self.addr5):
            if line:
                lines.append(line)
        city_line = ", ".join(p for p in (self.city, self.state) if p)
        if city_line and self.postal_code:
            city_line = f"{city_line} {self.postal_code}"
        elif self.postal_code:
            city_line = self.postal_code
        if city_line:
            lines.append(city_line)
        if self.country:
            lines.append(self.country)
        return "\n".join(lines)


class Ref(BaseEntity):
    list_id: Optional[str] = Field(default=None, alias="ListID")
    full_name: Optional[str] = Field(default=None, alias="FullName")
```

- [ ] **Step 4.5: Run — expect pass**

`pytest tests/unit/models -v`

- [ ] **Step 4.6: Commit**

```bash
git add src/qb_cli/models tests/unit/models
git commit -m "models: add BaseEntity, Address, Ref, money quantization"
```

---

## Task 5: QBXML envelope + common helpers

**Files:**
- Create: `src/qb_cli/qbxml/__init__.py` (empty)
- Create: `src/qb_cli/qbxml/envelope.py`
- Create: `src/qb_cli/qbxml/common.py`
- Test: `tests/unit/qbxml/test_envelope.py`
- Test: `tests/unit/qbxml/test_common.py`

- [ ] **Step 5.1: Write failing tests**

```python
# tests/unit/qbxml/test_envelope.py
from qb_cli.qbxml.envelope import wrap_request


def test_wrap_adds_qbxml_header_and_msgs():
    inner = '<InvoiceQueryRq requestID="1"></InvoiceQueryRq>'
    out = wrap_request(inner)
    assert out.startswith('<?xml')
    assert '<?qbxml version="16.0"?>' in out
    assert '<QBXMLMsgsRq onError="continueOnError">' in out
    assert inner in out
    assert out.strip().endswith("</QBXML>")
```

```python
# tests/unit/qbxml/test_common.py
from datetime import date
from qb_cli.qbxml.common import xml_escape, parse_address, format_qb_date


def test_xml_escape():
    assert xml_escape("A & B < C > 'x' \"y\"") == "A &amp; B &lt; C &gt; &apos;x&apos; &quot;y&quot;"


def test_format_qb_date():
    assert format_qb_date(date(2025, 5, 1)) == "2025-05-01"


def test_parse_address_from_element():
    from lxml import etree
    xml = b"""<ShipAddress>
        <Addr1>100 Main</Addr1><City>Seattle</City><State>WA</State>
        <PostalCode>98101</PostalCode>
    </ShipAddress>"""
    elem = etree.fromstring(xml)
    addr = parse_address(elem)
    assert addr.addr1 == "100 Main"
    assert addr.city == "Seattle"
```

- [ ] **Step 5.2: Run — expect fail**

- [ ] **Step 5.3: Implement `envelope.py`**

```python
# src/qb_cli/qbxml/envelope.py
def wrap_request(inner_body: str) -> str:
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        f'{inner_body}\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>\n'
    )
```

- [ ] **Step 5.4: Implement `common.py`**

```python
# src/qb_cli/qbxml/common.py
from __future__ import annotations

from datetime import date
from typing import Optional

from lxml import etree

from qb_cli.models.shared import Address


def xml_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
         .replace("<", "&lt;")
         .replace(">", "&gt;")
         .replace("'", "&apos;")
         .replace('"', "&quot;")
    )


def format_qb_date(d: date) -> str:
    return d.isoformat()


_ADDR_TAGS = ("Addr1", "Addr2", "Addr3", "Addr4", "Addr5",
              "City", "State", "PostalCode", "Country", "Note")


def parse_address(elem: Optional[etree._Element]) -> Optional[Address]:
    if elem is None:
        return None
    payload: dict[str, str] = {}
    for tag in _ADDR_TAGS:
        child = elem.find(tag)
        if child is not None and child.text:
            payload[tag] = child.text.strip()
    if not payload:
        return None
    return Address.model_validate(payload)
```

- [ ] **Step 5.5: Run — expect pass**

- [ ] **Step 5.6: Commit**

```bash
git add src/qb_cli/qbxml tests/unit/qbxml
git commit -m "qbxml: add envelope wrapper and common helpers"
```

---

## Task 6: Per-entity models ‖ (parallel fan-out, 15 subtasks)

Each subtask follows the same template. **An agent picks one entity, completes the subtask, and commits before picking the next (or another agent picks it up).**

### Template for `tasks/6.N/<entity>_model.md`

**Files:**
- Create: `src/qb_cli/models/<entity>.py`
- Test: `tests/unit/models/test_<entity>.py`

**Reference while coding:**
- QB field shapes: `/Users/lang/Documents/coding_temp/qb-mcp/src/qb_mcp/tools/<entity plural>/<entity>_add.py` (look at the XML fields assembled — those are the real field names).
- Do not copy code. Reference only.

**For each entity, the file contains:**

1. `class <Entity>LineItem(BaseEntity)` if the entity has line items. Fields common to most line items: `txn_line_id`, `item_ref` (Ref), `description`, `quantity` (Decimal), `rate` (Decimal), `amount` (Decimal), `customer_ref`, `class_ref`, `other1`, `other2`. Plus any entity-specific fields (see per-entity notes below).
2. `class <Entity>(BaseEntity)` with:
   - Identity: `txn_id` (alias `TxnID`, read-only for existing records), `edit_sequence` (alias `EditSequence`, read-only), `time_created`, `time_modified` (read-only).
   - Refs: `ref_number` (alias `RefNumber`), `txn_number` (alias `TxnNumber`, read-only).
   - Transaction-level: `txn_date`, `memo`, `customer_ref` or `vendor_ref`, `po_number`, `terms_ref`, `sales_rep_ref`, `ship_method_ref`, etc. (see per-entity notes).
   - Addresses: `bill_address`, `ship_address` (optional, both `Address`).
   - Line items: `line_items: list[<Entity>LineItem]` (if applicable).
   - Totals: `subtotal`, `sales_tax_total`, `total_amount`, `balance_remaining` (all `Decimal`, read-only where QB computes them).

3. `list[str] class-level constant REQUIRED_ON_ADD` — field names that must be present when building an Add request. Used by the validate path in the ops layer.

4. **Field-level constraint validators** (via `Field(max_length=...)` or `@field_validator`) that enforce QB's hard limits: RefNumber ≤ 11 chars, memo ≤ 4095, address lines ≤ 41 chars each, etc. Exact limits per entity are documented in the QBXML OSR (the linked reference in `qb-mcp/qb-docs/links-to-api-docs.txt`). When uncertain, copy from the matching `qb_mcp/tools/<entity>/<entity>_add.py` validation block.

**Template test file shape:**

```python
# tests/unit/models/test_<entity>.py
import pytest
from decimal import Decimal
from pydantic import ValidationError
from qb_cli.models.<entity> import <Entity>, <Entity>LineItem  # noqa
from qb_cli.models.shared import Address, Ref


def test_minimal_required_fields_accepted():
    # fill in with minimum set to create a valid record
    ...


def test_ref_number_max_length():
    with pytest.raises(ValidationError):
        <Entity>(ref_number="X" * 12)  # 12 chars > QB limit


def test_qb_alias_round_trip():
    # Build from QB-cased dict, dump with aliases — should re-match
    ...


def test_populates_line_items_from_qb_xml_casing():
    ...
```

**Steps (same for every entity):**

- [ ] **Step: Write failing test** (use template above; fill in entity specifics).
- [ ] **Step: Run — expect fail.**
- [ ] **Step: Implement model.**
- [ ] **Step: Run — expect pass** (`pytest tests/unit/models/test_<entity>.py -v`).
- [ ] **Step: Commit** (`git commit -m "models: add <entity> entity"`).

### Per-entity notes

- **6.1 invoice** — has `ship_date`, `po_number`, `is_pending`, `is_to_be_printed`, `is_to_be_emailed`, `customer_msg_ref`, `is_paid` (read-only), `applied_amount` (read-only). LineItem has `serial_number`, `lot_number`, `override_item_account_ref`, `is_taxable`.
- **6.2 sales_order** — has `ship_date`, `is_fully_invoiced` (read-only), `is_manually_closed`.
- **6.3 purchase_order** — `vendor_ref` (not customer), `vendor_address`, `expected_date`, `is_manually_closed`, `is_fully_received`.
- **6.4 bill** — two line-item containers: `expense_lines` (account-based) and `item_lines` (item-based). `vendor_ref`, `ap_account_ref`. Each expense line: `account_ref`, `amount`, `memo`, `customer_ref`, `class_ref`, `billable_status` (enum: Billable/NotBillable/HasBeenBilled).
- **6.5 estimate** — `is_active`, `markup`. Otherwise similar to invoice minus pending/paid.
- **6.6 sales_receipt** — `deposit_to_account_ref`, `payment_method_ref`, `check_number`. No `is_paid` (always paid by definition).
- **6.7 credit_memo** — `is_pending`, `is_to_be_printed`, credit-specific `is_auto_apply`.
- **6.8 receive_payment** — `customer_ref`, `ar_account_ref`, `payment_method_ref`, `total_amount`, `applied_to_txns: list[AppliedToTxn]` with each AppliedToTxn having `txn_id`, `payment_amount`, `discount_amount`, `discount_account_ref`.
- **6.9 check** — `account_ref` (bank), `payee_entity_ref` (customer/vendor/other_name), `is_to_be_printed`, `check_number`, `address`. Same `expense_lines`/`item_lines` split as Bill.
- **6.10 deposit** — `deposit_to_account_ref`, `cash_back_info`, `deposit_lines: list[DepositLine]` with `payment_txn_id` or `(entity_ref, account_ref, amount)`.
- **6.11 customer** — `name`, `company_name`, `first_name`, `middle_name`, `last_name`, `salutation`, `job_title`, `phone`, `alt_phone`, `fax`, `email`, `cc_email`, `contact`, `alt_contact`, `bill_address`, `ship_addresses: list[ShipAddress]` (note: list), `terms_ref`, `sales_rep_ref`, `tax_code_ref`, `price_level_ref`, `notes`, `credit_limit`, `balance` (read-only).
- **6.12 vendor** — `name`, `company_name`, `tax_id`, `is_vendor_eligible_for_1099`, `credit_limit`, `vendor_type_ref`, `terms_ref`, `bill_address`, `ship_address`, `contact_info` fields like customer.
- **6.13 item** — polymorphic. Define `ItemBase(BaseEntity)` with common fields (`list_id`, `name`, `full_name`, `is_active`, `parent_ref`, `sublevel`), then subclasses: `InventoryItem`, `ServiceItem`, `NonInventoryItem`, `OtherChargeItem`. In this task (6.13) only define the class hierarchy and (optionally) a `typing.Annotated` discriminated-union type `AnyItem = Annotated[Union[InventoryItem, ServiceItem, ...], Field(discriminator="item_type")]`. **The `parse_item(elem)` dispatcher lives in Task 7.13, not here** — do not block on it.
- **6.14 price_level** — `name`, `is_active`, `price_level_fixed_percentage` XOR `price_level_per_item: list[PriceLevelPerItem]` (with `item_ref`, `custom_price`).
- **6.15 ship_to** — `ship_to_address_block` fields (`name`, `addr1`..`addr5`, `city`, `state`, `postal_code`, `country`, `note`, `default_ship_to`). The `ship_to` entity is a sub-entity of `customer`; the model mirrors QB's `ShipToAddress` element.

**Agent dispatch note for tasks 6.1–6.15:** dispatch one subagent per entity after Task 5 is done. Each agent gets: the plan path, the spec path, the per-entity notes above, and the path to the matching `qb-mcp/src/qb_mcp/tools/<entity>*/` files as reference. Agents must not read each other's work or push to `main` — commits land on `main` serially once each subagent reports back and passes review.

---

## Task 7: Per-entity QBXML builders + parsers ‖ (parallel fan-out, 15 subtasks)

Depends on matching Task 6.N completing first.

### Template for `tasks/7.N/<entity>_qbxml.md`

**Files:**
- Create: `src/qb_cli/qbxml/<entity>.py`
- Test: `tests/unit/qbxml/test_<entity>_builder.py`
- Test: `tests/unit/qbxml/test_<entity>_parser.py`
- Fixture: `tests/fixtures/qbxml_responses/<entity>_query_sample.xml`

**Each module exports:**

```python
def build_query(
    *,
    ref_numbers: list[str] | None = None,
    txn_ids: list[str] | None = None,
    date_from: date | None = None,
    date_to: date | None = None,
    include_line_items: bool = True,
    max_returned: int | None = None,
    iterator_id: str | None = None,
) -> str: ...

def build_add(entity: <Entity>) -> str: ...
def build_mod(entity: <Entity>) -> str: ...

def parse_query_response(xml: str) -> list[<Entity>]: ...
def parse_add_response(xml: str) -> <Entity>: ...
def parse_mod_response(xml: str) -> <Entity>: ...
```

**Builder test shape (golden-file):**

```python
# tests/unit/qbxml/test_<entity>_builder.py
from datetime import date
from qb_cli.qbxml.<entity> import build_query


def test_build_query_by_ref_numbers():
    out = build_query(ref_numbers=["14396", "14397"], include_line_items=True)
    assert '<RefNumber>14396</RefNumber>' in out
    assert '<RefNumber>14397</RefNumber>' in out
    assert '<IncludeLineItems>true</IncludeLineItems>' in out
    assert '<?qbxml version="16.0"?>' in out


def test_build_query_date_range_order_valid():
    out = build_query(date_from=date(2025, 1, 1), date_to=date(2025, 12, 31))
    # TxnDateRangeFilter must come BEFORE IncludeLineItems (QBXML ordering rule)
    idx_filter = out.index('TxnDateRangeFilter')
    idx_include = out.index('IncludeLineItems')
    assert idx_filter < idx_include


def test_build_add_xml_escapes_memo():
    # Construct a minimal <Entity> with memo containing '&', assert escaped in output
    ...
```

**Parser test shape:**

```python
# tests/unit/qbxml/test_<entity>_parser.py
from pathlib import Path
from qb_cli.qbxml.<entity> import parse_query_response


FIXTURE = Path(__file__).parent.parent.parent / "fixtures" / "qbxml_responses" / "<entity>_query_sample.xml"


def test_parse_query_returns_expected_records():
    xml = FIXTURE.read_text()
    records = parse_query_response(xml)
    assert len(records) >= 1
    # Spot-check expected values (agent writing the fixture also writes these assertions)
    ...
```

**Fixture creation:** Each agent writes a minimal-but-realistic QBXML response fixture by hand. Reference samples live in `qb-mcp/src/qb_mcp/tools/<entity>*/<entity>_query.py` docstrings — copy shapes, not code. Alternatively, an agent can capture a real response from the running QB on the Windows host; either is fine as long as the fixture is checked in.

**Steps per entity (identical template):**

- [ ] **Step: Write failing builder tests.**
- [ ] **Step: Write fixture file + failing parser tests.**
- [ ] **Step: Run — expect fail.**
- [ ] **Step: Implement module.**
- [ ] **Step: Run — expect pass.**
- [ ] **Step: Commit** (`git commit -m "qbxml(<entity>): add builder and parser"`).

**Per-entity ordering notes (critical QBXML constraints):**

- Invoice/SO/Estimate/CreditMemo `*QueryRq` — `TxnDateRangeFilter` before `IncludeLineItems`.
- Add/Mod for transactions — line items come **after** all header fields and in the specific sub-element name QB expects (`InvoiceLineAdd` vs `InvoiceLineMod` vs `InvoiceLineRet`).
- `ItemQueryRq` returns a single response with mixed sub-elements (`ItemInventoryRet`, `ItemServiceRet`, ...); parser dispatches on element tag.
- `BillAddRq` requires at least one of `ExpenseLineAdd` or `ItemLineAdd`; builder enforces this precondition with a validation error, not a silent empty element.
- Mod requests require `EditSequence`; the builder raises a `ValueError` if the entity has no `edit_sequence` set.

---

## Task 8: IO serializers

### Task 8.1: Format enum

**Files:**
- Create: `src/qb_cli/io/__init__.py`
- Create: `src/qb_cli/io/format.py`
- Test: `tests/unit/io/test_format.py`

- [ ] **Step 8.1.1: Write failing tests**

```python
# tests/unit/io/test_format.py
import pytest
from qb_cli.io.format import Format, detect_format


def test_detect_by_extension():
    assert detect_format("x.csv") is Format.CSV
    assert detect_format("x.json") is Format.JSON
    assert detect_format("/path/y.JSON") is Format.JSON


def test_detect_unknown_raises():
    with pytest.raises(ValueError):
        detect_format("x.yaml")
```

- [ ] **Step 8.1.2: Run — expect fail.**

- [ ] **Step 8.1.3: Implement**

```python
# src/qb_cli/io/format.py
from __future__ import annotations

from enum import Enum
from pathlib import Path


class Format(str, Enum):
    CSV = "csv"
    JSON = "json"


def detect_format(path: str | Path) -> Format:
    ext = Path(path).suffix.lower().lstrip(".")
    try:
        return Format(ext)
    except ValueError as e:
        raise ValueError(f"unsupported file format '{ext}' for {path}") from e
```

- [ ] **Step 8.1.4: Run — expect pass. Commit.**

```bash
git add src/qb_cli/io tests/unit/io/test_format.py
git commit -m "io: add Format enum + detect_format"
```

### Task 8.2: JSON serializer

**Files:**
- Create: `src/qb_cli/io/json_serializer.py`
- Test: `tests/unit/io/test_json_roundtrip.py`

- [ ] **Step 8.2.1: Write failing roundtrip test**

```python
# tests/unit/io/test_json_roundtrip.py
from pathlib import Path
from qb_cli.models.invoice import Invoice
from qb_cli.io.json_serializer import to_json, from_json


def test_invoice_round_trip(tmp_path: Path):
    original = Invoice(
        ref_number="14396",
        customer_ref={"FullName": "ACME Corp"},
        txn_date="2025-05-01",
        line_items=[
            {"item_ref": {"FullName": "40-RAG12"}, "quantity": "10", "rate": "25.00", "amount": "250.00"},
        ],
    )
    path = tmp_path / "inv.json"
    to_json([original], path)
    loaded = from_json(Invoice, path)
    assert len(loaded) == 1
    assert loaded[0].ref_number == original.ref_number
    assert loaded[0].line_items[0].amount == original.line_items[0].amount
```

- [ ] **Step 8.2.2: Run — expect fail.**

- [ ] **Step 8.2.3: Implement**

```python
# src/qb_cli/io/json_serializer.py
from __future__ import annotations

import json
from pathlib import Path
from typing import Iterable, Type, TypeVar

from qb_cli.models.base import BaseEntity

T = TypeVar("T", bound=BaseEntity)


def to_json(entities: Iterable[BaseEntity], path: str | Path) -> None:
    records = [e.model_dump(by_alias=True, exclude_none=True, mode="json") for e in entities]
    Path(path).write_text(json.dumps(records, indent=2))


def from_json(model: Type[T], path: str | Path) -> list[T]:
    raw = json.loads(Path(path).read_text())
    return [model.model_validate(r) for r in raw]
```

- [ ] **Step 8.2.4: Run — expect pass. Commit.**

### Task 8.3: CSV serializer

**Files:**
- Create: `src/qb_cli/io/csv_serializer.py`
- Test: `tests/unit/io/test_csv_roundtrip.py`

**Behavior (per spec §5.6):**
- Output has a single header row = union of header-level + line-level columns.
- Every data row has `row_type` (`header`|`line`) and `parent_ref`.
- Flat entities (customers, vendors, etc.) emit only `header` rows.

- [ ] **Step 8.3.1: Write failing roundtrip tests**

```python
# tests/unit/io/test_csv_roundtrip.py
from pathlib import Path
from qb_cli.models.invoice import Invoice
from qb_cli.models.customer import Customer
from qb_cli.io.csv_serializer import to_csv, from_csv


def test_invoice_round_trip(tmp_path: Path):
    inv = Invoice(
        ref_number="14396",
        customer_ref={"FullName": "ACME Corp"},
        txn_date="2025-05-01",
        line_items=[
            {"item_ref": {"FullName": "40-RAG12"}, "quantity": "10", "rate": "25.00", "amount": "250.00"},
            {"item_ref": {"FullName": "Freight"}, "quantity": "1", "rate": "35.00", "amount": "35.00"},
        ],
    )
    path = tmp_path / "inv.csv"
    to_csv([inv], path)
    loaded = from_csv(Invoice, path)
    assert len(loaded) == 1
    assert len(loaded[0].line_items) == 2
    assert loaded[0].ref_number == "14396"
    assert loaded[0].line_items[0].amount == inv.line_items[0].amount


def test_flat_entity_no_line_rows(tmp_path: Path):
    c = Customer(name="ACME Corp", company_name="ACME Corporation, Inc.")
    path = tmp_path / "c.csv"
    to_csv([c], path)
    text = path.read_text()
    assert "row_type" in text
    assert "parent_ref" in text
    assert text.count("\n") == 2  # header row + 1 data row (+ trailing newline)
```

- [ ] **Step 8.3.2: Run — expect fail.**

- [ ] **Step 8.3.3: Implement**

Implementation strategy:
- `to_csv(entities, path)` — inspect the model class, derive header columns = `["row_type", "parent_ref"] + header_cols + line_cols_union`. For each entity, emit one header row; then for each list-typed field on the model (there may be more than one — e.g. `Bill` has both `expense_lines` and `item_lines`), emit one line row per item with `parent_ref` set to the entity's `ref_number`.
- `row_type` values: `header`, plus one value per list field (e.g. `line` for invoices/SOs/POs/estimates/credit_memos/sales_receipts, `expense_line` and `item_line` for bills, `deposit_line` for deposits, `applied_to` for receive_payments). The serializer derives these from the list field name (e.g. field `expense_lines` → row_type `expense_line`). Round-trip asserts must round through this.
- `from_csv(model, path)` — read all rows; group by `parent_ref`; bucket line rows by their `row_type` into the appropriate list field on the parent.
- Use Python's stdlib `csv` module. Preserve Decimal precision by emitting strings, not floats.
- Column discovery: `<Entity>.model_fields` gives header fields; filter list-typed fields whose element type is a `BaseEntity` subclass to discover line-item sub-field sets. The union of all such sub-fields' columns forms the line column block.

Full implementation lives in `src/qb_cli/io/csv_serializer.py`. Test `test_bill_round_trip_with_both_line_types` MUST be included — create a Bill with at least one expense line and one item line, serialize, reload, confirm both lists survive.

- [ ] **Step 8.3.4: Run — expect pass. Commit.**

```bash
git commit -m "io: add CSV serializer with round-trip support"
```

---

## Task 9: Ops layer

### Task 9.1: Entity registry

**Files:**
- Create: `src/qb_cli/ops/__init__.py`
- Create: `src/qb_cli/ops/registry.py`
- Test: `tests/unit/ops/test_registry.py`

Implement `@dataclass class EntityHandler` with: `key`, `model`, `qbxml_module`, `id_field` (e.g. `"ref_number"` for txns, `"name"` for list), `supports_mod`, `has_line_items`, `supports_dry_run`.

Populate `HANDLERS: dict[str, EntityHandler]` with all 15 entries.

Test: every entity key resolves; looking up a non-registered key raises `KeyError`.

- [ ] Steps: test → fail → impl → pass → commit.

### Task 9.2: Export op

**Files:**
- Create: `src/qb_cli/ops/export_op.py`
- Test: `tests/unit/ops/test_export_op.py`

```python
def export(
    ctx: Context,
    entity_key: str,
    *,
    ref_numbers: list[str] | None = None,
    date_from: date | None = None,
    date_to: date | None = None,
    output_path: Path,
    fmt: Format | None = None,  # None -> detect_format(output_path)
) -> ExportResult
```

Internally:
1. Look up handler in registry.
2. Call handler.qbxml_module.build_query(...).
3. Send via ctx.connection (mockable).
4. Parse response into list of model instances.
5. Serialize to output_path via io.
6. Return ExportResult with count, duration, output path.

Tests mock ctx.connection.send to return a canned XML fixture and verify the output file exists with expected rows.

- [ ] Steps: test → fail → impl → pass → commit.

### Task 9.3: Pre-flight verify helper

**Files:**
- Create: `src/qb_cli/ops/verify_op.py`
- Test: `tests/unit/ops/test_verify_op.py`

```python
def verify_entities(ctx, *, customers: list[str] = (), vendors: list[str] = (),
                   items: list[str] = (), terms: list[str] = (), accounts: list[str] = (),
                   ship_methods: list[str] = (), sales_reps: list[str] = ()) -> VerifyResult
```

For each non-empty list, issue one `*QueryRq` filtered by names; compare returned names against input; return `VerifyResult(found=..., missing=...)`.

- [ ] Steps: test → fail → impl → pass → commit.

### Task 9.4: Duplicate detection

**Files:**
- Create: `src/qb_cli/ops/dedupe.py`
- Test: `tests/unit/ops/test_dedupe.py`

```python
def find_duplicates(ctx, entity_key: str, ref_numbers: list[str]) -> set[str]
```

Issue a batched `<Entity>QueryRq` by ref numbers; any that come back exist. Return the set of already-existing refs.

- [ ] Steps: test → fail → impl → pass → commit.

### Task 9.5: Import op

**Files:**
- Create: `src/qb_cli/ops/import_op.py`
- Test: `tests/unit/ops/test_import_op.py`

```python
def import_(
    ctx: Context,
    entity_key: str,
    *,
    input_path: Path,
    fmt: Format | None = None,
    dry_run: bool = False,
    on_duplicate: Literal["error", "skip", "update"] = "error",
) -> ImportResult
```

Flow:
1. Read entities from file via io.
2. Collect all referenced customers/vendors/items/terms; call `verify_entities`. If anything missing → abort with `ImportResult.missing`.
3. Collect all entity ref_numbers; call `find_duplicates`.
4. Decide action for each duplicate per `on_duplicate`:
   - `error` → abort with list of duplicates.
   - `skip` → remove from to-write set.
   - `update` → fetch edit_sequence via query_op, convert to Mod.
5. If `dry_run`: return ImportResult with plan only.
6. Otherwise: build + send Add/Mod requests one at a time (batching is optional v1.5). Collect successes and errors. Return ImportResult.

Tests cover each branch with mocked ctx.connection.

- [ ] Steps: test → fail → impl → pass → commit.

### Task 9.6: Query op

**Files:**
- Create: `src/qb_cli/ops/query_op.py`
- Test: `tests/unit/ops/test_query_op.py`

Thin wrapper: given `entity_key` and refs/filters, return list of model instances. Used both by CLI `query` command and internally by `import_op`'s duplicate/update flow.

- [ ] Steps: test → fail → impl → pass → commit.

### Task 9.7: Context object + utils

**Files:**
- Create: `src/qb_cli/context.py`
- Create: `src/qb_cli/config.py`
- Create: `src/qb_cli/utils/__init__.py`
- Create: `src/qb_cli/utils/logging.py`
- Create: `src/qb_cli/utils/date_parse.py`
- Test: `tests/unit/test_context.py`
- Test: `tests/unit/utils/test_logging.py`
- Test: `tests/unit/utils/test_date_parse.py`

```python
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Callable, ContextManager

ConnectionFactory = Callable[[], ContextManager["QBConnection"]]


@dataclass
class Context:
    connection_factory: ConnectionFactory
    config: Config
    logger: logging.Logger
    output_dir: Path
    default_format: Format
    json_output: bool
    dry_run: bool
```

`Context` holds a **connection factory**, not a live connection, because the REPL stays open across many commands and a single QB session per-op is safer (§5.7 of the spec). Ops call `with ctx.connection_factory() as conn: ...`. The factory is built from config + CLI overrides.

`utils/logging.py` — `setup_logging(level: str, json: bool) -> logging.Logger` returning a configured root logger. JSON mode emits one-line JSON records; text mode emits a short human-readable format.

`utils/date_parse.py` — `parse_date(s: str) -> date` accepts `YYYY-MM-DD`, `MM/DD/YYYY`, and `today`/`yesterday` shortcuts; raises `ValueError` otherwise. `year_range(year: int) -> tuple[date, date]` returns `(YYYY-01-01, today if year==current_year else YYYY-12-31)`.

`Config` loaded from `~/.config/qb_cli/config.toml` (or `%APPDATA%\qb_cli\config.toml` on Windows) with env-var overrides. `Context.from_env(**cli_overrides)` assembles the object.

- [ ] Steps for each file: test → fail → impl → pass → commit (4 commits total: logging, date_parse, config, context).

---

## Task 10: CLI (click command tree)

**Files:**
- Create: `src/qb_cli/cli/__init__.py`, `cli/root.py`, `cli/export_cmd.py`, `cli/import_cmd.py`, `cli/query_cmd.py`, `cli/add_cmd.py`, `cli/mod_cmd.py`, `cli/verify_cmd.py`, `cli/repl_cmd.py`
- Test: `tests/unit/cli/test_*.py` (one per command)

### Task 10.1: Root group

```python
# src/qb_cli/cli/root.py
import click
from qb_cli.context import Context
from qb_cli.cli.export_cmd import export
from qb_cli.cli.import_cmd import import_
from qb_cli.cli.query_cmd import query
from qb_cli.cli.add_cmd import add
from qb_cli.cli.mod_cmd import mod
from qb_cli.cli.verify_cmd import verify
from qb_cli.cli.repl_cmd import repl


@click.group(invoke_without_command=True)
@click.option("--company-file", default=None, envvar="QB_CLI_COMPANY_FILE")
@click.option("--config", "config_path", type=click.Path())
@click.option("--log-level", default="INFO", envvar="QB_CLI_LOG_LEVEL")
@click.option("--json", "json_output", is_flag=True, help="emit machine-readable JSON")
@click.pass_context
def qb(ctx, company_file, config_path, log_level, json_output):
    ctx.obj = Context.from_env(
        company_file=company_file,
        config_path=config_path,
        log_level=log_level,
        json_output=json_output,
    )
    if ctx.invoked_subcommand is None:
        ctx.invoke(repl)


qb.add_command(export)
qb.add_command(import_, name="import")
qb.add_command(query)
qb.add_command(add)
qb.add_command(mod)
qb.add_command(verify)
qb.add_command(repl)
```

- [ ] Steps: test (`CliRunner` + `--help` assertion) → fail → impl → pass → commit.

### Task 10.2–10.7: One command per file

Each command takes the entity key as first argument, looks up the handler in the registry, and dispatches to the corresponding op. `--dry-run` belongs on `import`, `add`, `mod` only.

Template for `export_cmd.py`:

```python
import click
from pathlib import Path
from qb_cli.ops.export_op import export
from qb_cli.ops.registry import HANDLERS
from qb_cli.io.format import Format, detect_format


@click.command("export")
@click.argument("entity", type=click.Choice(list(HANDLERS.keys())))
@click.option("--ref", "ref_numbers", multiple=True)
@click.option("--year", type=int, help="export all records from Jan 1 of YEAR through today")
@click.option("--from", "date_from", type=click.DateTime(formats=["%Y-%m-%d"]))
@click.option("--to", "date_to", type=click.DateTime(formats=["%Y-%m-%d"]))
@click.option("--out", "output_path", type=click.Path(dir_okay=False, writable=True), required=True)
@click.option("--format", "fmt", type=click.Choice(["csv", "json"]), default=None)
@click.pass_obj
def export_cmd(ctx, entity, ref_numbers, year, date_from, date_to, output_path, fmt):
    ...
```

- [ ] Per-command steps: test → fail → impl → pass → commit (6 commits, one per file).

### Task 10.8: Wire `__main__`

Replace placeholder with:

```python
# src/qb_cli/__main__.py
from qb_cli.cli.root import qb as main


if __name__ == "__main__":
    main()
```

- [ ] Run `qb --help` and `qb export --help` manually; confirm. Commit.

---

## Task 11: REPL

### Task 11.1: Parser (line → argv)

**Files:**
- Create: `src/qb_cli/repl/__init__.py`, `repl/parser.py`
- Test: `tests/unit/repl/test_parser.py`

`parse_line(line: str) -> list[str]` — shlex-based. Reject empty, strip comments, handle trailing backslash continuations (v1: no, just split).

- [ ] Steps: test → fail → impl → pass → commit.

### Task 11.2: Completer

**Files:**
- Create: `src/qb_cli/repl/completers.py`
- Test: `tests/unit/repl/test_completers.py`

Provide a prompt_toolkit `Completer` subclass that reads the current Click group and emits sub-command completions plus entity keys from the registry. Field-level completion (v2) is out of scope for v1.

- [ ] Steps: test → fail → impl → pass → commit.

### Task 11.3: Shell loop

**Files:**
- Create: `src/qb_cli/repl/shell.py`
- Test: `tests/unit/repl/test_shell.py`

Use `prompt_toolkit.PromptSession`:

```python
def run_repl(ctx: Context) -> int:
    session = PromptSession(
        message=_build_prompt(ctx),
        history=FileHistory(_history_path()),
        completer=QBCompleter(),
    )
    while True:
        try:
            line = session.prompt()
        except (EOFError, KeyboardInterrupt):
            return 0
        argv = parse_line(line)
        if not argv:
            continue
        if argv[0] in {"exit", "quit"}:
            return 0
        # dispatch to click
        try:
            from qb_cli.cli.root import qb
            qb.main(args=argv, standalone_mode=False, obj=ctx)
        except click.ClickException as e:
            e.show()
        except Exception as e:
            click.echo(f"error: {e}", err=True)
```

Tests drive with a scripted input stream via prompt_toolkit's `create_pipe_input` helper. See prompt_toolkit docs (use `context7` if unsure of the current API).

- [ ] Steps: test → fail → impl → pass → commit.

---

## Task 12: Integration smoke suite (live QB)

**Files:**
- Create: `tests/integration/test_smoke.py`
- Create: `tests/integration/conftest.py`

Require `--run-integration` flag; gate with `pytest.mark.integration`. One test per entity, each one runs export with a short date range and expects ≥ 0 records (zero is a valid result on a fresh company file).

```python
# tests/integration/conftest.py
import pytest


def pytest_collection_modifyitems(config, items):
    if not config.getoption("--run-integration", default=False):
        skip_int = pytest.mark.skip(reason="--run-integration not passed")
        for item in items:
            if "integration" in item.keywords:
                item.add_marker(skip_int)


def pytest_addoption(parser):
    parser.addoption("--run-integration", action="store_true", default=False)
```

Agents cannot run these (no Windows QB in dev env). Include them; user runs them on Windows. Commit.

---

## Task 13: README rewrite

**Files:**
- Modify: root `README.md`

Write from scratch covering:
- What qb-cli is
- Install (with the `[windows]` extra on the QB machine)
- `qb` REPL quickstart
- Non-interactive CLI examples (export/import/query/verify) for one representative entity
- Table of supported entity types and which commands work per entity
- Pointer to spec + plan docs
- Pointer to `archive/` for legacy scripts

- [ ] Steps: draft → commit (`docs: rewrite README for qb-cli tool`).

---

## Done conditions

Task is complete when:
- All 13 top-level tasks are committed on `main`.
- `pytest -m "not integration"` passes on macOS (no pywin32, no live QB) with 100% of non-integration tests green.
- On the Windows QB host, `pip install -e '.[dev,windows]'` succeeds, and `qb export invoice --year 2025 --out /tmp/inv.csv` produces a non-empty file for a company with 2025 invoices.
- `qb` with no args launches the REPL; tab-completion works for verbs and entity keys; `exit` quits cleanly.
- `archive/` contains all six original scripts unchanged; the repo's git history preserves them via `git mv`.

---

## Dispatch notes for subagent-driven-development

Suggested wave structure for the execution skill:

- **Wave 0:** Task 1 (scaffold) → Task 2 (archive) → Task 3 (transport) → Task 4 (model base) → Task 5 (qbxml common). Sequential. Single agent chain.
- **Wave 1:** Tasks 6.1–6.15 in parallel (15 agents). Each delivers one model + tests + commit.
- **Wave 2:** Tasks 7.1–7.15 in parallel (15 agents), each depending on its matching 6.N.
- **Wave 3:** Tasks 8.1, 8.2, 8.3 in parallel (3 agents) — can run concurrently with Wave 2.
- **Wave 4:** Tasks 9.1 → 9.2/9.3/9.4/9.6 parallel → 9.5 (blocks on 9.3+9.4) → 9.7.
- **Wave 5:** Task 10 sequentially (10.1 → 10.2..10.7 parallel → 10.8).
- **Wave 6:** Task 11 sequentially.
- **Wave 7:** Tasks 12 + 13 in parallel.

Each subagent's prompt must include:
- The path to this plan and the spec.
- Its task number (e.g. 6.7 credit_memo).
- Its task's files-to-create and test template.
- The "read qb-mcp as a reference, do not copy code" rule.
- The type-hint rule (no `Any`).
- Instruction to commit only after tests pass, one commit per task.
- Prohibition on pushing to remote; `main` is local-only until the user asks.
