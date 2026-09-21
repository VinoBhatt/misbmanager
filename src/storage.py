"""D1 workbook and record storage, with SQLite for local development."""
import base64
import hashlib
import io
import os
from pathlib import Path
from flask import request, g

# Encoded chunks stay comfortably below D1's 2 MB row limit.
CHUNK_BYTES = 512 * 1024
MAX_WORKBOOK_BYTES = 10 * 1024 * 1024


def workbook_statements(key, data):
    if not data or len(data) > MAX_WORKBOOK_BYTES:
        raise ValueError('Workbook must contain data and be at most 10 MB')
    chunks = [data[i:i+CHUNK_BYTES] for i in range(0, len(data), CHUNK_BYTES)]
    statements = [("INSERT INTO workbook_versions(object_key,byte_size,chunk_count,sha256) VALUES(?,?,?,?)",
                   (key, len(data), len(chunks), hashlib.sha256(data).hexdigest()))]
    statements.extend(("INSERT INTO workbook_chunks(object_key,chunk_index,content) VALUES(?,?,?)",
                       (key, i, base64.b64encode(chunk).decode('ascii'))) for i, chunk in enumerate(chunks))
    return statements


def atomic_write(statements):
    if cloud():
        from pyodide.ffi import run_sync
        db = cloud().DB
        # D1 batch rolls back the complete batch if any statement fails.
        run_sync(db.batch([db.prepare(sql).bind(*params) for sql, params in statements]))
    else:
        con = connect()
        try:
            with con:
                for sql, params in statements:
                    con.execute(sql, params)
        finally:
            con.close()


def cloud():
    return request.environ.get("workers.env")


def local_dir():
    return Path(os.environ.get("MISB_DATA_DIR", ".local"))


class D1Cursor:
    def __init__(self, result):
        self.rows = result.get("results", [])
        self.lastrowid = result.get("meta", {}).get("last_row_id")

    def __iter__(self):
        return iter(self.rows)


class D1Connection:
    def execute(self, sql, params=()):
        from pyodide.ffi import run_sync
        stmt = cloud().DB.prepare(sql)
        if params:
            stmt = stmt.bind(*params)
        return D1Cursor(run_sync(stmt.all()))

    def commit(self):
        pass  # Each D1 statement commits atomically.

    def close(self):
        pass


def connect():
    if cloud():
        return D1Connection()
    import sqlite3
    con = sqlite3.connect(local_dir() / "misb_tracker.db")
    con.row_factory = sqlite3.Row
    return con


def source_rows():
    if "sources" not in g:
        con = connect()
        try:
            g.sources = {r["kind"]: dict(r) for r in con.execute("SELECT * FROM sources")}
        finally:
            con.close()
    return g.sources


def source_exists(kind):
    return kind in source_rows()


def read_source(kind):
    if "source_bytes" not in g:
        g.source_bytes = {}
    if kind not in g.source_bytes:
        row = source_rows().get(kind)
        if not row:
            raise ValueError(f"Import the {kind} workbook from Data Sources first.")
        con = connect()
        try:
            versions = list(con.execute('SELECT * FROM workbook_versions WHERE object_key=?', (row['object_key'],)))
            if not versions:
                raise ValueError(f"Import the {kind} workbook into the database from Data Sources.")
            version = versions[0]
            chunks = list(con.execute('SELECT chunk_index,content FROM workbook_chunks WHERE object_key=? ORDER BY chunk_index', (row['object_key'],)))
        finally:
            con.close()
        if len(chunks) != version['chunk_count'] or any(r['chunk_index'] != i for i, r in enumerate(chunks)):
            raise ValueError('Stored workbook is incomplete')
        data = b''.join(base64.b64decode(r['content'], validate=True) for r in chunks)
        if len(data) != version['byte_size'] or hashlib.sha256(data).hexdigest() != version['sha256']:
            raise ValueError('Stored workbook failed its integrity check')
        g.source_bytes[kind] = data
    return io.BytesIO(g.source_bytes[kind])


def save_source(kind, data, as_of):
    # Publish the complete workbook and its pointer in one transaction.
    from uuid import uuid4
    key = f"sources/{kind}/{uuid4().hex}.xlsx"
    previous = source_rows().get(kind, {}).get("object_key")
    statements = workbook_statements(key, data)
    statements.append(("""INSERT INTO sources(kind,object_key,as_of) VALUES(?,?,?)
          ON CONFLICT(kind) DO UPDATE SET object_key=excluded.object_key,
          as_of=excluded.as_of,updated_at=CURRENT_TIMESTAMP""", (kind, key, as_of)))
    atomic_write(statements)
    g.pop("sources", None)
    g.pop("source_bytes", None)
    return previous
