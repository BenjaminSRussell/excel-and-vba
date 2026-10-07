# SQLite sidecar schema

Pair every `.xlsm` with an adjacent `.db` for audit/recovery.

## Tables
```sql
CREATE TABLE runs (
  id INTEGER PRIMARY KEY,
  started_at TEXT NOT NULL,
  finished_at TEXT,
  note TEXT
);
CREATE TABLE queue_items (
  id INTEGER PRIMARY KEY,
  run_id INTEGER REFERENCES runs(id),
  payload TEXT NOT NULL,
  status TEXT NOT NULL,
  created_at TEXT NOT NULL,
  updated_at TEXT
);
CREATE TABLE params (
  id INTEGER PRIMARY KEY,
  run_id INTEGER REFERENCES runs(id),
  key TEXT NOT NULL,
  value TEXT NOT NULL
);
```

Backup both the workbook and the `.db`.
