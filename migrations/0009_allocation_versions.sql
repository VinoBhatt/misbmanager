CREATE TABLE IF NOT EXISTS note_allocation_versions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  loan_code TEXT NOT NULL,
  action TEXT NOT NULL,
  changed_fields TEXT NOT NULL DEFAULT '[]',
  snapshot TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_note_allocation_versions_code
ON note_allocation_versions(loan_code, id DESC);
