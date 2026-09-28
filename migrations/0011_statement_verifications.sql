CREATE TABLE IF NOT EXISTS statement_report_verifications (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  statement_run_id INTEGER NOT NULL,
  matches INTEGER NOT NULL,
  mismatches TEXT NOT NULL DEFAULT '[]',
  verified_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_statement_report_verifications_run
ON statement_report_verifications(statement_run_id, id DESC);
