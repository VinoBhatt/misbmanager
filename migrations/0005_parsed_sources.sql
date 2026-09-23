CREATE TABLE IF NOT EXISTS parsed_sources (
  kind TEXT NOT NULL,
  object_key TEXT NOT NULL,
  payload TEXT NOT NULL,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (kind, object_key)
);
