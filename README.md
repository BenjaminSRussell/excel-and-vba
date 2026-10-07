# excel-and-vba

Excel/VBA calculation and queue tooling with an optional SQLite audit sidecar.

## Layout
- `src/modules/*.bas` — exported VBA modules (text)
- `workbooks/` — workbook specs (binary `.xlsm` stays local; see docs)
- `docs/` — queue workbook + storage guide
- `scripts/queue_sidecar.py` — Python helper that writes audit rows

## Quick start
```bash
python scripts/queue_sidecar.py --db workbooks/queue_audit.db --enqueue "demo-job"
sqlite3 workbooks/queue_audit.db "SELECT * FROM queue_items;"
```
