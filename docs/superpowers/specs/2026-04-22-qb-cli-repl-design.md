# qb-cli: Interactive REPL + Non-interactive CLI for QuickBooks Desktop

**Date:** 2026-04-22
**Status:** Draft — awaiting user review
**Author:** brainstorming session

---

## 1. Summary

Build a single Python tool that replaces the current folder of standalone
QuickBooks export scripts with one program that exposes every operation through
two entry points sharing the same grammar:

1. A **non-interactive CLI** with verb-noun subcommands
   (`qb export invoice --year 2025 --out invoices.csv`).
2. An **interactive REPL** with history, tab completion, and stateful session
   context (`qb repl` → `qb> export invoice --year 2025`).

The tool lives in the existing `qb-export-scripts` repository. Existing
standalone scripts are moved to `archive/` (preserved, unmodified) while the new
tool is developed under `src/qb_cli/`. Once the new tool has feature parity
with the old scripts, the archive can be deleted in a later change.

The tool supports both **export** (CSV or JSON output) and, for the first time,
**import** (CSV or JSON input, creating or modifying QuickBooks records).
Exports and imports are round-trippable: exporting a record, then re-importing
the produced file, is a no-op assuming no other changes in QuickBooks.

## 2. Goals and Non-Goals

### Goals

- One binary (`qb`) that functions as both CLI and REPL.
- Identical command grammar between CLI and REPL so that scripts and interactive
  use share muscle memory.
- Coverage in v1 for **core transactions** and **core list entities** (see §4).
- Round-trippable **CSV + JSON** I/O for every supported entity.
- **Pre-flight validation** (verify referenced entities exist) and
  **duplicate detection** (by `RefNumber`) before any write.
- Fresh QBXML implementation authored in this project; `qb-mcp` is read as a
  reference for correct XML shapes, but its code is not imported, vendored, or
  installed as a dependency.

### Non-Goals (v1)

- **Not** an MCP server or MCP client. This tool never speaks the MCP protocol.
- **Not** a full parity port of `qb-mcp`'s ~236 tools.
- **Not** cross-platform at runtime. QuickBooks Desktop requires Windows and
  pywin32; the CLI is developed on macOS but only runs against a live QB on
  Windows.
- **Not** dry-run-by-default. Imports write by default; `--dry-run` is an
  opt-in flag. Pre-flight verify and duplicate detection are the primary safety
  nets.

## 3. User Stories

1. As a user running on the Windows machine where QB is installed, I can type
   `qb` and land in a REPL with history, tab completion, and a persistent
   session so I can iterate on exports and imports without restarting.
2. As the same user, I can invoke the same commands non-interactively
   (`qb export invoice --year 2025 --out invoices.csv`) from a shell script or
   scheduled task.
3. I can export invoices to CSV, edit the file in Excel, and import the edited
   file back to apply modifications to the original invoices.
4. I can take a JSON dump of 50 bills from a staging environment and import
   them into production; before anything is written, I get told if any
   referenced vendors, items, or accounts are missing, and if any of the
   incoming RefNumbers already exist.
5. I can add new customers in bulk from a CSV file.
6. I can ask what data exists (`qb query customer "ACME Corp"`) without leaving
   the REPL.

## 4. Entity Scope (v1)

### Transactions (full CRUD: query / add / mod / export / import)

| Entity | QBXML request root | Existing script? |
|---|---|---|
| `invoice` | `InvoiceQueryRq`, `InvoiceAddRq`, `InvoiceModRq` | yes (`qb_inv.py`) |
| `sales_order` | `SalesOrderQueryRq`, … | yes (`qb_so.py`) |
| `purchase_order` | `PurchaseOrderQueryRq`, … | yes (`qb_po.py`) |
| `bill` | `BillQueryRq`, … | yes (`qb_bills.py`) |
| `estimate` | `EstimateQueryRq`, … | no |
| `sales_receipt` | `SalesReceiptQueryRq`, … | no |
| `credit_memo` | `CreditMemoQueryRq`, … | no |
| `receive_payment` | `ReceivePaymentQueryRq`, … | no |
| `check` | `CheckQueryRq`, … | no |
| `deposit` | `DepositQueryRq`, … | no |

### Core lists (query / add / mod / export / import)

| Entity | QBXML request root | Existing script? |
|---|---|---|
| `customer` | `CustomerQueryRq`, `CustomerAddRq`, `CustomerModRq` | partial (`example_customer_query.py`) |
| `vendor` | `VendorQueryRq`, … | no |
| `item` | `ItemQueryRq` (fans out to Item*AddRq per subtype) | no |

### Carry-overs from existing scripts

| Entity | Notes |
|---|---|
| `price_level` | Already covered by `qb_price_level.py`; port export + add new import. |
| `ship_to` | Already covered by `qb_shipto.py`; included for parity. |

**Total v1 entities: 15.**

For item subtypes (inventory, service, non-inventory, other-charge), the
external command is `qb add item` with an `--item-type` flag; internally this
dispatches to the correct `Item*AddRq`.

Each entity supports the operations below (list entities skip operations that
QuickBooks itself does not support, e.g. hard delete):

- `export <entity>` — pull from QB, write CSV or JSON
- `import <entity>` — read CSV or JSON, add or modify in QB
- `query <entity>` — quick read to stdout (no file I/O)
- `add <entity>` — create a single record from command-line flags
- `mod <entity>` — modify a single record from command-line flags
- `verify <entity>` — check existence / validity without writing

## 5. Architecture

### 5.1 Layered structure

```
           ┌────────────────────┐
           │  repl (prompt_tk)  │
           └─────────┬──────────┘
                     │  reuses parser
           ┌─────────▼──────────┐
           │   cli (click)      │
           └─────────┬──────────┘
                     ▼
           ┌────────────────────┐
           │   ops              │  export / import / query / verify / dedupe
           └─────┬──────┬───────┘
                 │      │
     ┌───────────▼┐   ┌─▼────────────┐
     │ io         │   │ models       │   pydantic entities
     │ csv, json  │   └─▲────────────┘
     └───────────┬┘     │
                 │      │
           ┌─────▼──────▼───────┐
           │ qbxml              │   per-entity builders + parsers
           └─────────┬──────────┘
                     ▼
           ┌────────────────────┐
           │ transport          │   pywin32 COM + QBXML wire
           └────────────────────┘
```

Every arrow points in exactly one direction; lower layers never import higher
layers.

### 5.2 Module layout

```
qb-export-scripts/
├── archive/                         # old scripts, unchanged
│   ├── qb_inv.py
│   ├── qb_so.py
│   ├── qb_po.py
│   ├── qb_bills.py
│   ├── qb_price_level.py
│   ├── qb_shipto.py
│   ├── example_customer_query.py
│   ├── example_purchase_order_add.py
│   ├── test_ship_to_query.py
│   └── README-archived.md
├── docs/
│   └── superpowers/specs/2026-04-22-qb-cli-repl-design.md  (this file)
├── src/qb_cli/
│   ├── __init__.py
│   ├── __main__.py                  # `python -m qb_cli` entry
│   ├── version.py
│   ├── config.py                    # TOML config loader
│   ├── context.py                   # Runtime Context object
│   ├── transport/
│   │   ├── __init__.py
│   │   ├── connection.py            # COM RequestProcessor wrapper
│   │   ├── session.py               # OpenConnection/BeginSession context mgr
│   │   ├── errors.py                # Exception hierarchy
│   │   └── qbxml_io.py              # send_qbxml(request) -> response
│   ├── qbxml/
│   │   ├── __init__.py
│   │   ├── envelope.py              # QBXML header/envelope helpers
│   │   ├── common.py                # Address parsers, ref parsers, date fmt
│   │   ├── invoice.py
│   │   ├── sales_order.py
│   │   ├── purchase_order.py
│   │   ├── bill.py
│   │   ├── estimate.py
│   │   ├── sales_receipt.py
│   │   ├── credit_memo.py
│   │   ├── receive_payment.py
│   │   ├── check.py
│   │   ├── deposit.py
│   │   ├── customer.py
│   │   ├── vendor.py
│   │   ├── item.py
│   │   ├── price_level.py
│   │   └── ship_to.py
│   ├── models/
│   │   ├── __init__.py
│   │   ├── base.py                  # BaseEntity (pydantic BaseModel subclass)
│   │   ├── shared.py                # Address, LineItem, Ref etc.
│   │   ├── invoice.py ... (one per entity, matching qbxml/)
│   ├── io/
│   │   ├── __init__.py
│   │   ├── csv_serializer.py
│   │   ├── json_serializer.py
│   │   └── format.py                # Format enum + file extension dispatch
│   ├── ops/
│   │   ├── __init__.py
│   │   ├── export_op.py
│   │   ├── import_op.py
│   │   ├── query_op.py
│   │   ├── verify_op.py
│   │   ├── dedupe.py
│   │   └── registry.py              # entity_type -> handler map
│   ├── cli/
│   │   ├── __init__.py
│   │   ├── root.py                  # click.Group for `qb`
│   │   ├── export_cmd.py
│   │   ├── import_cmd.py
│   │   ├── query_cmd.py
│   │   ├── add_cmd.py
│   │   ├── mod_cmd.py
│   │   ├── verify_cmd.py
│   │   └── repl_cmd.py
│   ├── repl/
│   │   ├── __init__.py
│   │   ├── shell.py                 # prompt_toolkit main loop
│   │   ├── completers.py            # Custom completers
│   │   └── parser.py                # Line -> argv -> click
│   └── utils/
│       ├── __init__.py
│       ├── logging.py
│       └── date_parse.py
├── tests/
│   ├── __init__.py
│   ├── unit/
│   │   ├── qbxml/
│   │   │   ├── test_invoice_builder.py
│   │   │   ├── test_invoice_parser.py
│   │   │   └── ... (one per entity)
│   │   ├── io/
│   │   │   ├── test_csv_roundtrip.py
│   │   │   └── test_json_roundtrip.py
│   │   ├── models/
│   │   ├── ops/
│   │   └── cli/
│   ├── integration/                 # require live QB, marker-gated
│   └── fixtures/
│       ├── qbxml_responses/         # sample XML from real QB
│       └── golden/                  # expected serialized outputs
├── pyproject.toml
├── README.md                        # rewritten for new tool
└── requirements.txt                 # deprecated; kept for backward compat
```

### 5.3 Transport layer

- `transport/connection.py` wraps
  `win32com.client.Dispatch('QBXMLRP2.RequestProcessor')`. It exposes a
  `QBConnection` class that is a context manager handling
  `OpenConnection2`/`BeginSession`/`EndSession`/`CloseConnection`.
  The connection mode defaults to `QBFileMode.qbFileOpenDoNotCare` (2).
- `transport/qbxml_io.py` exposes `send(connection, qbxml_request: str) -> str`.
  No entity-level semantics live here.
- `transport/errors.py` defines:
  - `QBError` (base)
  - `QBConnectionError`
  - `QBXMLParseError`
  - `QBStatusError` (raised when the response `statusCode != 0`; carries
    `status_code`, `status_message`, `status_severity`, `request_id`)
  - `QBEntityNotFound` (special case of `QBStatusError` for status 500/3120)
  - `QBDuplicateEntity` (special case for status 3100/3270)

### 5.4 QBXML layer

One module per entity. Each module defines pure functions — no I/O, no COM
calls. Signatures follow a consistent shape:

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

def build_add(entity: InvoiceModel) -> str: ...
def build_mod(entity: InvoiceModel) -> str: ...

def parse_query_response(xml: str) -> list[InvoiceModel]: ...
def parse_add_response(xml: str) -> InvoiceModel: ...
def parse_mod_response(xml: str) -> InvoiceModel: ...
```

Transactions that exceed QB's per-request limit use QBXML iterators; the
`build_query` helpers expose the iterator parameters so the ops layer can pump
pages without the qbxml layer caring about I/O.

Request-order constraints imposed by QBXML (filter elements before
`IncludeLineItems`, etc.) are enforced here. This is the layer that encodes
QBXML's irregularities; nothing above it should know.

XML construction uses `lxml.etree` (cleaner namespaces, avoids string
concatenation mistakes). Parsing uses `lxml.etree` with explicit element
visitors per entity — no generic schema reflection.

### 5.5 Models

- `BaseEntity(pydantic.BaseModel)` with `model_config = ConfigDict(populate_by_name=True, extra='forbid')`.
- Fields use Python-friendly names (`ref_number`) with `alias='RefNumber'` so
  JSON serialization matches QB casing.
- Shared submodels in `models/shared.py`: `Address`, `LineItem`, `Ref`
  (which QB calls `ListID/FullName` pairs), `Amount` (Decimal with 2-dp quantize).
- Read-only fields (e.g. `TxnID`, `EditSequence`, `TimeCreated`) use
  `Field(frozen=True)`; they're populated by QBXML parsing, never sent back
  except as part of a Mod request where `EditSequence` is required.
- Pydantic validators enforce QB constraints (RefNumber ≤ 11 chars, memo ≤ 4095
  chars, address lines ≤ 5, etc.).

### 5.6 IO serializers

#### JSON

Straight Pydantic `model_dump_json`. Round-trip invariant is trivial.

#### CSV

A single header row that is the **union** of header-level and line-level
columns. Every data row includes a `row_type` column (`header` or `line`) and a
`parent_ref` column that pins line rows to their header row by `ref_number`.
Columns that do not apply to a row type are left blank.

Concrete example for `invoice` (abridged — real schema has more columns):

```
row_type,parent_ref,ref_number,customer,txn_date,po_number,line_item,quantity,rate,amount
header,,14396,ACME Corp,2025-05-01,7740-SH,,,,
line,14396,,,,,40-RAG12,10,25.00,250.00
line,14396,,,,,Freight,1,35.00,35.00
header,,14397,Other Co,2025-05-02,7741,,,,
line,14397,,,,,24-CW412P,100,0.50,50.00
```

Flat entities (customers, vendors, simple items, price levels, ship-tos) write
only `row_type=header` rows; `parent_ref` and the line columns stay empty.

`from_csv` groups consecutive rows by `parent_ref` to reassemble line items.
Round-trip: `from_csv(to_csv(xs)) == xs`, property-tested.

### 5.7 Ops layer

Each op takes a `Context` (carries connection factory, config, logger, dry-run
flag) plus entity-specific arguments, and returns a `Result` dataclass with
`success_count`, `failure_count`, `records`, and `errors` fields.

- `export_op(ctx, entity_type: str, filters, output_path, fmt)` — calls
  qbxml.build_query, sends via transport, parses response, pages as needed,
  writes via io.
- `import_op(ctx, entity_type: str, input_path, fmt, dry_run, on_duplicate)` —
  reads via io, runs verify, runs dedupe, sends add/mod requests, reports.
- `query_op(ctx, entity_type, refs)` — thin wrapper, prints to stdout.
- `verify_op(ctx, entity_type, refs)` — checks existence, returns presence map.
- `dedupe.check_duplicates(ctx, entity_type, ref_numbers) -> set[str]` —
  returns refs that already exist.

The entity-to-handler lookup lives in `ops/registry.py`:

```python
HANDLERS: dict[str, EntityHandler] = {
    "invoice": EntityHandler(
        model=InvoiceModel,
        qbxml_module=qb_cli.qbxml.invoice,
        id_field="ref_number",
        supports_mod=True,
    ),
    # ...
}
```

### 5.8 CLI

- Built on **Click** (mature, composable, plays well with prompt_toolkit).
- Root group `qb` carries global flags: `--company-file`, `--config`,
  `--log-level`, `--json`. `--dry-run` is a flag on `import`, `add`, and `mod`
  subcommands only (it has no meaning on read operations).
- Subcommand shape: `qb <verb> <noun> [options]`.
- `--help` pages auto-generated. A `--help-entities` root flag lists all
  supported entity types.

Example invocations:

```bash
qb export invoice --year 2025 --out invoices.csv
qb export invoice --ref 14396 --ref 14397 --out inv.json --format json
qb import bill --file vendors.csv --on-duplicate skip
qb query customer "ACME Corp"
qb verify item 40-RAG12 Freight "Net 30"
qb add customer --name "New Co" --company "New Company, Inc."
qb mod invoice --ref 14396 --po-number 7740-SH
```

### 5.9 REPL

- Launched by `qb repl`, or by `qb` with no arguments.
- Uses **prompt_toolkit** for the input loop.
- Prompt is `qb> ` by default; includes `[dry]` when dry-run is active and
  `[disconnected]` when no QB connection is live.
- History persisted to `~/.local/state/qb_cli/history` (or platform
  equivalent).
- Tab completion via a custom `Completer`:
  - At column 0: complete verbs (`export`, `import`, …).
  - After a verb: complete entity types from `ops/registry.py`.
  - After a flag: entity-specific value completers (file paths, years, refs).
- REPL-only meta commands:
  - `set <key> <value>` / `unset <key>` — mutate Context (dry_run, out_dir, format).
  - `status` — print session state.
  - `connect` / `disconnect` — explicit COM session control for long REPL sessions.
  - `help [verb]` — mirror `qb --help`.
  - `quit` / `exit` / Ctrl-D — exit.
- Command dispatch: split line with `shlex.split`, prepend to click's
  `standalone_mode=False` invocation of the root group. Exceptions are caught
  and pretty-printed; the REPL does not exit on errors.

### 5.10 Config and Context

- Config lives at `~/.config/qb_cli/config.toml` (XDG on macOS/Linux, appdata
  on Windows).
- Keys: `company_file`, `default_output_dir`, `default_format`, `log_level`.
- Env overrides: `QB_CLI_COMPANY_FILE`, etc.
- Command-line flags override env which overrides config.
- Context object is assembled in the root Click callback and stashed on
  `click.Context.obj`; REPL loop reuses the same object across commands.

### 5.11 Safety features

- **Pre-flight verify** is wired into `import_op` unconditionally. Before any
  write, every referenced customer, vendor, item, account, and terms value is
  checked (one batched `*Query` request per entity type). If anything is
  missing, the op aborts with a structured list of missing refs and zero
  writes.
- **Duplicate detection** runs after pre-flight. For each incoming
  `ref_number`, a `*Query` checks if it already exists. Behavior controlled by
  `--on-duplicate`:
  - `error` (default) — abort the whole import on the first duplicate.
  - `skip` — skip the duplicates, write the rest.
  - `update` — convert each duplicate from an add into a mod (requires
    fetching `EditSequence` first).
- **Dry-run** (`--dry-run`) runs pre-flight and dedupe but stops before any
  `Add`/`Mod` request. Prints the plan.

### 5.12 Error handling

- All QB errors surface as `QBError` subclasses with `status_code`,
  `status_message`, `request_id`.
- The CLI top-level catches `QBError`, prints a human-readable one-line
  summary plus the original QB status message, exits with a non-zero code.
- The REPL catches the same, prints, continues.
- Logging: `qb_cli.utils.logging` configures a `structlog`-style logger
  (plain `logging` with a JSON formatter when `--json` is set, text otherwise).

### 5.13 Testing strategy

- **Unit tests** dominate. Every `qbxml/<entity>.py` module has a golden-file
  test: build a request from fixed inputs, compare against a canonical XML
  snapshot; parse a fixed response XML, assert the model matches.
- **Serializer round-trip tests**: property tests using `hypothesis` to fuzz
  `from_csv(to_csv(x)) == x` and `from_json(to_json(x)) == x` for every
  entity.
- **Ops tests** mock the transport layer (fake `send()` that returns stubbed
  XML) to exercise pagination, pre-flight, dedupe, and error paths.
- **CLI tests** use `click.testing.CliRunner`.
- **REPL tests** drive the shell with a scripted input stream.
- **Integration tests** (marked `pytest.mark.integration`, skipped by default)
  require live QB and exercise a small smoke suite per entity against a test
  company file.

### 5.14 Dependencies

| Package | Purpose | Notes |
|---|---|---|
| `click` | CLI framework | |
| `prompt_toolkit` | REPL input loop | |
| `pydantic` | Entity models | v2 |
| `lxml` | XML build + parse | already used by `qb-mcp` |
| `pywin32` | QB COM | Windows-only runtime dep |
| `pytest` | Tests | dev |
| `hypothesis` | Property tests | dev |
| `tomli` | TOML config parse on py3.10 | hard dep at floor; code uses `tomllib` via `try/except ImportError` shim |

Python `>=3.10` to match `qb-mcp`. On 3.10, `tomli` provides `tomllib` semantics;
on 3.11+ the stdlib `tomllib` is used.

## 6. Build Sequence (handoff to writing-plans)

Ordered; items within a step marked `‖` can run as parallel agent tasks.

1. **Scaffold** — `pyproject.toml`, `src/qb_cli/` skeleton, `tests/` skeleton,
   archive move of existing scripts, updated README stub.
2. **Transport layer** — connection, session, errors, `send` (one agent; small).
3. **Model foundation** — `models/base.py`, `models/shared.py`.
4. ‖ **Per-entity models** (15 entities) — each entity is one file, no
   cross-dependencies after step 3.
5. ‖ **Per-entity QBXML builders + parsers** (15 entities) — depend on step 4.
6. ‖ **IO serializers** — CSV + JSON, parallel with step 5.
7. **Ops layer** — `export_op`, `import_op`, `query_op`, `verify_op`, `dedupe`,
   `registry`. Depends on 4–6.
8. **CLI** — click command tree. Depends on 7.
9. **REPL** — prompt_toolkit shell. Depends on 8.
10. **Docs + README rewrite**.
11. **Smoke suite** — one integration test per entity against a live QB
    (deferred to user; agents produce the scaffolding).

## 7. Resolved decisions and open questions

### Resolved in this spec

- **`--on-duplicate update` is in v1.** It reuses the `mod` path (fetches
  `EditSequence` via a query, then issues a `Mod` request). Deferring it would
  buy very little since the machinery already exists for `mod`.
- **Item import uses a single CSV with an `item_type` column** (values:
  `inventory`, `service`, `non_inventory`, `other_charge`). The ops layer
  dispatches to the correct `Item*AddRq` based on this column. The CLI flag for
  a single-record add is `--item-type` (kebab-case per Click convention); the
  CSV column is `item_type` (snake-case per data convention). This mapping is
  the standard translation and needs no special code.

### Still open, non-blocking

- Exact default location of the company file on the Windows host — the config
  reads from `config.toml` with no hard-coded default. Defaults are fine as-is
  for planning.
