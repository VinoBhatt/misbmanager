CREATE TABLE IF NOT EXISTS allocation_email_tracking (
    loan_code TEXT PRIMARY KEY,
    allocation_email_date TEXT NOT NULL DEFAULT '',
    disbursement_email_date TEXT NOT NULL DEFAULT '',
    ledger_signature TEXT NOT NULL,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
