ALTER TABLE note_allocations ADD COLUMN approval_status TEXT NOT NULL DEFAULT 'Draft';
ALTER TABLE note_allocations ADD COLUMN approval_updated_at TEXT;

CREATE TABLE IF NOT EXISTS source_version_meta (
  object_key TEXT PRIMARY KEY REFERENCES workbook_versions(object_key),
  kind TEXT NOT NULL,
  as_of TEXT NOT NULL DEFAULT '',
  filename TEXT NOT NULL DEFAULT ''
);

INSERT OR IGNORE INTO source_version_meta(object_key,kind,as_of)
SELECT object_key,kind,COALESCE(as_of,'') FROM sources;
