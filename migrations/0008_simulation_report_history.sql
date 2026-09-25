CREATE TABLE IF NOT EXISTS simulation_report_runs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  as_of TEXT NOT NULL,
  filename TEXT NOT NULL,
  baseline_rows INTEGER NOT NULL DEFAULT 0,
  new_rows INTEGER NOT NULL DEFAULT 0,
  ledger_updated_rows INTEGER NOT NULL DEFAULT 0,
  total_rows INTEGER NOT NULL DEFAULT 0,
  missing_allocations INTEGER NOT NULL DEFAULT 0,
  snapshot TEXT NOT NULL DEFAULT '{}',
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_simulation_report_runs_created
ON simulation_report_runs(created_at DESC, id DESC);
