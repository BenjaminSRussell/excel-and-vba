#!/usr/bin/env python3
"""Write queue audit rows to the SQLite sidecar."""
from __future__ import annotations
import argparse
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS runs (
  id INTEGER PRIMARY KEY,
  started_at TEXT NOT NULL,
  finished_at TEXT,
  note TEXT
);
CREATE TABLE IF NOT EXISTS queue_items (
  id INTEGER PRIMARY KEY,
  run_id INTEGER REFERENCES runs(id),
  payload TEXT NOT NULL,
  status TEXT NOT NULL,
  created_at TEXT NOT NULL,
  updated_at TEXT
);
CREATE TABLE IF NOT EXISTS params (
  id INTEGER PRIMARY KEY,
  run_id INTEGER REFERENCES runs(id),
  key TEXT NOT NULL,
  value TEXT NOT NULL
);
"""

def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default="workbooks/queue_audit.db")
    ap.add_argument("--enqueue", required=True)
    ap.add_argument("--note", default="python-helper")
    args = ap.parse_args()

    path = Path(args.db)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.executescript(SCHEMA)
    now = datetime.now(timezone.utc).isoformat()
    cur = conn.execute("INSERT INTO runs (started_at, note) VALUES (?, ?)", (now, args.note))
    run_id = cur.lastrowid
    conn.execute(
        "INSERT INTO queue_items (run_id, payload, status, created_at) VALUES (?, ?, 'pending', ?)",
        (run_id, args.enqueue, now),
    )
    conn.commit()
    print({"db": str(path), "run_id": run_id, "payload": args.enqueue})
    conn.close()

if __name__ == "__main__":
    main()
