"""D1 workbook and record storage, with SQLite for local development."""
import base64
import hashlib
import io
import json
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


def _read_workbook_bytes(object_key):
    con = connect()
    try:
        versions = list(con.execute('SELECT * FROM workbook_versions WHERE object_key=?', (object_key,)))
        if not versions: raise ValueError('That workbook version was not found.')
        version = versions[0]
        chunks = list(con.execute('SELECT chunk_index,content FROM workbook_chunks WHERE object_key=? ORDER BY chunk_index', (object_key,)))
    finally:
        con.close()
    if len(chunks) != version['chunk_count'] or any(r['chunk_index'] != i for i, r in enumerate(chunks)):
        raise ValueError('Stored workbook is incomplete')
    data = b''.join(base64.b64decode(r['content'], validate=True) for r in chunks)
    if len(data) != version['byte_size'] or hashlib.sha256(data).hexdigest() != version['sha256']:
        raise ValueError('Stored workbook failed its integrity check')
    return data


def read_source_version(kind, object_key):
    """Read an immutable workbook after confirming it belongs to this source."""
    versions={row['object_key']:row for row in source_versions(kind)}
    if object_key not in versions:
        raise ValueError('That workbook version does not belong to this source.')
    return io.BytesIO(_read_workbook_bytes(object_key)),versions[object_key]


def read_source(kind):
    if "source_bytes" not in g:
        g.source_bytes = {}
    if kind not in g.source_bytes:
        row = source_rows().get(kind)
        if not row:
            raise ValueError(f"Import the {kind} workbook from Data Sources first.")
        try: g.source_bytes[kind] = _read_workbook_bytes(row['object_key'])
        except ValueError as error:
            if 'not found' in str(error): raise ValueError(f"Import the {kind} workbook into the database from Data Sources.")
            raise
    return io.BytesIO(g.source_bytes[kind])


def read_parsed_source(kind):
    """Return normalized data for the currently selected workbook version."""
    row = source_rows().get(kind)
    if not row:
        return None
    con = connect()
    try:
        matches = list(con.execute(
            'SELECT payload FROM parsed_sources WHERE kind=? AND object_key=?',
            (kind, row['object_key'])
        ))
    finally:
        con.close()
    return json.loads(matches[0]['payload']) if matches else None


def save_parsed_source(kind, object_key, payload):
    """Store JSON-safe normalized data beside its immutable workbook version."""
    encoded = json.dumps(payload, separators=(',', ':'), default=lambda value: value.isoformat())
    con = connect()
    try:
        con.execute('''INSERT INTO parsed_sources(kind,object_key,payload) VALUES(?,?,?)
          ON CONFLICT(kind,object_key) DO UPDATE SET payload=excluded.payload,
          updated_at=CURRENT_TIMESTAMP''', (kind, object_key, encoded))
        con.commit()
    finally:
        con.close()


def save_source(kind, data, as_of, filename=''):
    # Publish the complete workbook and its pointer in one transaction.
    from uuid import uuid4
    key = f"sources/{kind}/{uuid4().hex}.xlsx"
    previous_row = source_rows().get(kind, {})
    previous = previous_row.get("object_key")
    statements = workbook_statements(key, data)
    if previous:
        statements.append(("INSERT OR IGNORE INTO source_version_meta(object_key,kind,as_of) VALUES(?,?,?)",
                           (previous, kind, previous_row.get('as_of',''))))
    statements.append(("INSERT INTO source_version_meta(object_key,kind,as_of,filename) VALUES(?,?,?,?)",
                       (key, kind, as_of, str(filename or ''))))
    statements.append(("""INSERT INTO sources(kind,object_key,as_of) VALUES(?,?,?)
          ON CONFLICT(kind) DO UPDATE SET object_key=excluded.object_key,
          as_of=excluded.as_of,updated_at=CURRENT_TIMESTAMP""", (kind, key, as_of)))
    atomic_write(statements)
    g.pop("sources", None)
    g.pop("source_bytes", None)
    return previous


def source_versions(kind):
    """List stored versions for one source, newest first."""
    current=source_rows().get(kind,{}).get('object_key')
    con=connect()
    try:
        rows=list(con.execute('''SELECT v.object_key,v.byte_size,v.sha256,v.created_at,
          COALESCE(m.as_of,'') AS as_of,COALESCE(m.filename,'') AS filename FROM workbook_versions v
          LEFT JOIN source_version_meta m ON m.object_key=v.object_key
          WHERE m.kind=? OR v.object_key=? OR v.object_key LIKE ?
          ORDER BY v.created_at DESC''',(kind,current or '',f'sources/{kind}/%')))
    finally:
        con.close()
    return [{**dict(row),'current':row['object_key']==current} for row in rows]


def select_source_version(kind, object_key, as_of=''):
    """Point a source at an existing immutable workbook version."""
    versions={row['object_key']:row for row in source_versions(kind)}
    if object_key not in versions:
        raise ValueError('That workbook version does not belong to this source.')
    chosen_as_of=as_of or versions[object_key].get('as_of') or source_rows().get(kind,{}).get('as_of','')
    con=connect()
    try:
        con.execute('UPDATE sources SET object_key=?,as_of=?,updated_at=CURRENT_TIMESTAMP WHERE kind=?',
                    (object_key,chosen_as_of,kind))
        con.commit()
    finally:
        con.close()
    g.pop('sources',None);g.pop('source_bytes',None)
    return chosen_as_of


def audit_event(action, entity_type, entity_id='', details=None):
    """Append a compact immutable management event to D1."""
    con=connect()
    try:
        con.execute('INSERT INTO audit_events(action,entity_type,entity_id,details) VALUES(?,?,?,?)',
                    (action,entity_type,str(entity_id or ''),json.dumps(details or {},separators=(',',':'))))
        con.commit()
    finally:
        con.close()


def audit_rows(limit=200):
    con=connect()
    try:
        rows=[dict(row) for row in con.execute('SELECT * FROM audit_events ORDER BY id DESC LIMIT ?',
                                               (max(1,min(int(limit),500)),))]
    finally:
        con.close()
    for row in rows:
        try: row['details']=json.loads(row.get('details') or '{}')
        except (TypeError,json.JSONDecodeError): row['details']={}
    return rows


def save_simulation_report_run(as_of, filename, metrics, snapshot):
    con=connect()
    try:
        con.execute('''INSERT INTO simulation_report_runs(
            as_of,filename,baseline_rows,new_rows,ledger_updated_rows,total_rows,missing_allocations,snapshot
        ) VALUES(?,?,?,?,?,?,?,?)''',(
            as_of,filename,int(metrics.get('baseline_rows',0)),int(metrics.get('new_rows',0)),
            int(metrics.get('ledger_updated_rows',0)),int(metrics.get('total_rows',0)),
            int(metrics.get('missing_allocations',0)),json.dumps(snapshot,separators=(',',':'))
        ))
        con.commit()
    finally:
        con.close()


def simulation_report_runs(limit=20):
    con=connect()
    try:
        rows=[dict(row) for row in con.execute('''SELECT id,as_of,filename,baseline_rows,new_rows,
            ledger_updated_rows,total_rows,missing_allocations,snapshot,created_at
            FROM simulation_report_runs ORDER BY id DESC LIMIT ?''',(max(1,min(int(limit),100)),))]
    finally:
        con.close()
    for row in rows:
        try: row['snapshot']=json.loads(row.get('snapshot') or '{}')
        except (TypeError,json.JSONDecodeError): row['snapshot']={}
    return rows


def simulation_report_run(run_id):
    con=connect()
    try:
        rows=[dict(row) for row in con.execute('''SELECT id,as_of,filename,baseline_rows,new_rows,
            ledger_updated_rows,total_rows,missing_allocations,snapshot,created_at
            FROM simulation_report_runs WHERE id=?''',(int(run_id),))]
    finally:
        con.close()
    if not rows: return None
    result=rows[0]
    try: result['snapshot']=json.loads(result.get('snapshot') or '{}')
    except (TypeError,json.JSONDecodeError): result['snapshot']={}
    return result


def save_statement_report_run(as_of, source_object_key, filename, metrics):
    con=connect()
    try:
        con.execute('''INSERT INTO statement_report_runs(
            as_of,source_object_key,filename,row_count,page_count,opening_balance,closing_balance,gross_returns
        ) VALUES(?,?,?,?,?,?,?,?)''',(
            as_of,source_object_key,filename,int(metrics.get('row_count',0)),int(metrics.get('page_count',0)),
            float(metrics.get('opening_balance',0)),float(metrics.get('closing_balance',0)),
            float(metrics.get('gross_returns',0))))
        con.commit()
    finally:
        con.close()


def statement_report_runs(limit=20):
    con=connect()
    try:
        return [dict(row) for row in con.execute('''SELECT id,as_of,source_object_key,filename,row_count,
            page_count,opening_balance,closing_balance,gross_returns,created_at
            FROM statement_report_runs ORDER BY id DESC LIMIT ?''',(max(1,min(int(limit),100)),))]
    finally:
        con.close()


def statement_report_run(run_id):
    con=connect()
    try:
        rows=[dict(row) for row in con.execute('''SELECT id,as_of,source_object_key,filename,row_count,
            page_count,opening_balance,closing_balance,gross_returns,created_at
            FROM statement_report_runs WHERE id=?''',(int(run_id),))]
    finally:
        con.close()
    return rows[0] if rows else None
