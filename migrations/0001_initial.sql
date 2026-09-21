CREATE TABLE IF NOT EXISTS monitoring (
 loan_code TEXT PRIMARY KEY, status TEXT DEFAULT 'Normal', owner TEXT DEFAULT '',
 next_action TEXT DEFAULT '', review_date TEXT DEFAULT '', notes TEXT DEFAULT '',
 updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
INSERT OR IGNORE INTO settings VALUES ('issuer_limit', '2000000');
CREATE TABLE IF NOT EXISTS payment_marks (
 loan_code TEXT NOT NULL, event_key TEXT NOT NULL, event_type TEXT NOT NULL,
 label TEXT DEFAULT '', expected_date TEXT DEFAULT '', expected_amount REAL DEFAULT 0,
 is_paid INTEGER DEFAULT 0, paid_date TEXT DEFAULT '', paid_amount REAL DEFAULT 0,
 notes TEXT DEFAULT '', updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
 PRIMARY KEY (loan_code,event_key)
);
CREATE TABLE IF NOT EXISTS cashflow_plan (
 id INTEGER PRIMARY KEY AUTOINCREMENT, flow_date TEXT NOT NULL, flow_type TEXT NOT NULL,
 loan_code TEXT DEFAULT '', issuer TEXT DEFAULT '', amount REAL NOT NULL,
 notes TEXT DEFAULT '', created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS sources (
 kind TEXT PRIMARY KEY, object_key TEXT NOT NULL, as_of TEXT DEFAULT '',
 updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
