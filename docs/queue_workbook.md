# Sample FIFO queue workbook

## Sheets
- `Queue` columns: `id`, `payload`, `status` (`pending`/`done`), `created_at`
- `Params` named ranges for calc inputs

## Manual test cases
1. Enqueue three rows via `QueueModule.Enqueue`.
2. Run `QueueModule.ProcessNext` twice; expect two `done`, one `pending`.
3. Confirm SQLite sidecar gains matching `queue_items` rows (see `docs/storage.md`).
