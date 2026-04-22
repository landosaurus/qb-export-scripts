# Archived standalone scripts

These are the original one-off QuickBooks export scripts. They are preserved
here as a reference and for emergency use while the new `qb` CLI is under
development. The new tool lives in `src/qb_cli/` and is documented in the
root `README.md`.

Each script is standalone — it issues its own QBXML and writes a CSV.
They continue to work on the Windows QuickBooks machine as long as pywin32
is installed.

Do not add new features here. New work belongs in `src/qb_cli/`.
