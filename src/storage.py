"""Request-scoped Cloudflare bindings, with a SQLite/filesystem local adapter."""
import io
import os
from pathlib import Path
from flask import request, g


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
        if cloud():
            from pyodide.ffi import run_sync
            obj = run_sync(cloud().FILES.get(row["object_key"]))
            if obj is None:
                raise ValueError(f"The {kind} workbook is missing from storage.")
            data = bytes(run_sync(obj.arrayBuffer()))
        else:
            data = (local_dir() / row["object_key"]).read_bytes()
        g.source_bytes[kind] = data
    return io.BytesIO(g.source_bytes[kind])


def save_source(kind, data, as_of):
    # Write immutable objects first. A single atomic pointer update publishes
    # the validated upload, leaving earlier versions available for recovery.
    from uuid import uuid4
    key = f"sources/{kind}/{uuid4().hex}.xlsx"
    if cloud():
        from pyodide.ffi import run_sync
        run_sync(cloud().FILES.put(key, data))
    else:
        target = local_dir() / key
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
    previous = source_rows().get(kind, {}).get("object_key")
    con = connect()
    try:
        con.execute("""INSERT INTO sources(kind,object_key,as_of) VALUES(?,?,?)
          ON CONFLICT(kind) DO UPDATE SET object_key=excluded.object_key,
          as_of=excluded.as_of,updated_at=CURRENT_TIMESTAMP""", (kind, key, as_of))
        con.commit()
    finally:
        con.close()
    g.pop("sources", None)
    g.pop("source_bytes", None)
    return previous
