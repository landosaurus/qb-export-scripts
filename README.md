# qb-cli

Interactive REPL, guided wizard, and non-interactive CLI for QuickBooks Desktop.

`qb-cli` exports, imports, queries, and verifies records (invoices, sales orders, purchase orders) against a local QuickBooks Desktop company file via QBXML over the COM bridge.

> **Status:** under active development. Live QuickBooks operations require Windows + QuickBooks Desktop + `pywin32`. Everything else (parsing, serialization, dry-run flows) is cross-platform.

## Install

```bash
# from a checkout
pip install -e .

# on the QuickBooks Windows machine
pip install -e ".[windows]"
```

This installs a `qb` console script.

## Three ways to use it

### 1. Wizard (default)

Running `qb` with no arguments launches an interactive menu-driven wizard. It walks you through entity choice, date range / ref selection, output format, and confirmation. Useful for one-off jobs and for users who don't want to memorize flags.

```
$ qb
? What would you like to do?
  > Export records
    Import records
    Query records
    Verify entities exist
    Exit
```

Use `b` / `m` at any text prompt (or pick the `← Back` / `← Main menu` choices) to navigate.

### 2. REPL

`qb repl` opens a persistent shell with tab-completion, command history, and the same sub-commands as the CLI. Useful when you're running several commands against the same company file and don't want to pay the startup / connection cost each time.

```
$ qb repl
qb> export invoice --year 2025 --out invoices_2025.csv
exported 412 invoice records to invoices_2025.csv
qb> query sales_order SO-1001 SO-1002
SO-1001  2025-08-14  1,200.00
SO-1002  2025-08-15    875.50
qb> exit
```

History is stored at `~/.local/state/qb_cli/history` (or `%APPDATA%\qb_cli\history` on Windows).

### 3. Direct CLI

Every wizard / REPL action is also a flat sub-command, suitable for scripts and cron jobs.

```bash
qb export invoice --year 2025 --out invoices_2025.csv
qb export purchase_order --from 2025-01-01 --to 2025-03-31 --out po_q1.json
qb query sales_order SO-1001 SO-1002
qb import invoice --file invoices_to_load.csv --dry-run
qb import invoice --file invoices_to_load.csv --on-duplicate skip
qb verify --customer "Acme Corp" --item "WIDGET-01"
```

## Commands

| Command   | What it does                                                                  |
| --------- | ----------------------------------------------------------------------------- |
| `export`  | Pull records from QB to CSV or JSON. Select by `--ref`, `--year`, or `--from/--to`. |
| `import`  | Push records from a CSV/JSON file into QB. Supports `--dry-run` and `--on-duplicate {error,skip,update}`. |
| `query`   | Print a one-line summary per record to stdout (ref, date, total).              |
| `verify`  | Check that referenced customers / vendors / items / terms exist in QB.         |
| `repl`    | Launch the interactive REPL.                                                  |
| `wizard`  | Launch the interactive wizard (also the default when `qb` is run with no args). |

Supported entities: `invoice`, `sales_order`, `purchase_order`.

Run `qb <command> --help` for the full flag list.

## Configuration

Settings resolve in this order (later wins): defaults → TOML config → environment variables → CLI flags.

**TOML** at `~/.config/qb_cli/config.toml` (or `%APPDATA%\qb_cli\config.toml` on Windows):

```toml
company_file = "C:\\Users\\Public\\Documents\\Intuit\\QuickBooks\\Company Files\\MyCompany.QBW"
default_output_dir = "C:\\qb-exports"
default_format = "csv"
log_level = "INFO"
```

**Environment variables:**

| Variable                       | Purpose                              |
| ------------------------------ | ------------------------------------ |
| `QB_CLI_COMPANY_FILE`          | Path to the `.QBW` company file.     |
| `QB_CLI_DEFAULT_OUTPUT_DIR`    | Default directory for exports.       |
| `QB_CLI_DEFAULT_FORMAT`        | `csv` or `json`.                     |
| `QB_CLI_LOG_LEVEL`             | `DEBUG`, `INFO`, `WARNING`, …        |

**Global flags** (apply to every sub-command): `--company-file`, `--config`, `--log-level`, `--json` (single-line JSON logs).

## Project layout

```
src/qb_cli/
  cli/        # click sub-commands
  repl/       # prompt-toolkit shell, parser, completer, wizard
  ops/        # export / import / query / verify operations
  qbxml/      # QBXML envelope + per-entity builders/parsers
  models/     # pydantic entity models
  io/         # CSV / JSON serializers, format detection
  transport/  # QB COM connection + error hierarchy
  config.py   # TOML + env config loader
  context.py  # runtime Context plumbed through every op
```

Design and implementation notes live under `docs/superpowers/`.

## Archived scripts

The original one-off export scripts are preserved under `archive/` for reference. New work belongs in `src/qb_cli/`; see `archive/README-archived.md`.
