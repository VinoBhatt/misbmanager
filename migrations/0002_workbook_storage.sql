CREATE TABLE IF NOT EXISTS workbook_versions (
 object_key TEXT PRIMARY KEY,
 byte_size INTEGER NOT NULL,
 chunk_count INTEGER NOT NULL,
 sha256 TEXT NOT NULL,
 created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS workbook_chunks (
 object_key TEXT NOT NULL REFERENCES workbook_versions(object_key),
 chunk_index INTEGER NOT NULL,
 content TEXT NOT NULL,
 PRIMARY KEY(object_key, chunk_index)
);
