# workbooks/

Local working directory for Excel binaries and the SQLite audit sidecar.

- `*.xlsm` / `*.xlsb` workbooks stay on your machine and are **never committed** (see `.gitignore`).
  Rebuild one from the exported modules in `../src/modules/` by following `../docs/queue_workbook.md`.
- `queue_audit.db` is created here the first time you run `scripts/queue_sidecar.py`.
  It is also gitignored (`*.db`).

Only this README is tracked, so a fresh clone matches the layout documented in the root README.
