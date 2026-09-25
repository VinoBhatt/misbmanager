CREATE TABLE IF NOT EXISTS statement_report_runs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  as_of TEXT NOT NULL,
  source_object_key TEXT NOT NULL,
  filename TEXT NOT NULL,
  row_count INTEGER NOT NULL DEFAULT 0,
  page_count INTEGER NOT NULL DEFAULT 0,
  opening_balance REAL NOT NULL DEFAULT 0,
  closing_balance REAL NOT NULL DEFAULT 0,
  gross_returns REAL NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_statement_report_runs_created
ON statement_report_runs(created_at DESC, id DESC);
