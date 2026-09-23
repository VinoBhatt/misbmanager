"""Run the site on loopback with local persistent storage for development."""
import argparse
import os
from pathlib import Path
import sqlite3
import sys

ROOT=Path(__file__).resolve().parent.parent
sys.path.insert(0,str(ROOT/'src'))


def initialize(directory, seed=None):
    directory.mkdir(parents=True,exist_ok=True)
    con=sqlite3.connect(directory/'misb_tracker.db')
    con.execute('CREATE TABLE IF NOT EXISTS local_migrations(name TEXT PRIMARY KEY, applied_at TEXT DEFAULT CURRENT_TIMESTAMP)')
    for migration in sorted((ROOT/'migrations').glob('*.sql')):
        if con.execute('SELECT 1 FROM local_migrations WHERE name=?',(migration.name,)).fetchone():
            continue
        con.executescript(migration.read_text())
        con.execute('INSERT INTO local_migrations(name) VALUES(?)',(migration.name,))
    from storage import workbook_statements
    # Carry existing local workbook files into SQLite without changing pointers.
    for key, in con.execute('SELECT object_key FROM sources').fetchall():
        if not con.execute('SELECT 1 FROM workbook_versions WHERE object_key=?',(key,)).fetchone():
            path=(directory/key).resolve()
            if path.is_relative_to(directory.resolve()) and path.is_file():
                for sql, params in workbook_statements(key,path.read_bytes()):
                    con.execute(sql,params)
    if seed:
        import json
        mapping={'transactions':'transactions.xlsx','simulation':'simulation.xlsx','projections':'midas_projections.xlsx'}
        for kind,filename in mapping.items():
            if con.execute('SELECT 1 FROM sources WHERE kind=?',(kind,)).fetchone():
                continue
            source=seed/filename
            if source.exists():
                key='sources/initial/'+filename
                for sql, params in workbook_statements(key,source.read_bytes()):
                    con.execute(sql,params)
                meta=seed/'source_meta.json'
                as_of=json.loads(meta.read_text()).get('simulation_as_of','') if meta.exists() else ''
                con.execute('INSERT INTO sources(kind,object_key,as_of) VALUES(?,?,?)',(kind,key,as_of))
        original=seed/'misb_tracker.db'
        if original.exists():
            con.execute('ATTACH DATABASE ? AS legacy',(str(original),))
            for table in ('monitoring','settings','payment_marks','cashflow_plan'):
                exists=con.execute("SELECT 1 FROM legacy.sqlite_master WHERE type='table' AND name=?",(table,)).fetchone()
                if exists:
                    con.execute(f'INSERT OR IGNORE INTO main.{table} SELECT * FROM legacy.{table}')
    con.commit();con.close()


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--seed',type=Path,help='Import original local data once (never overwrite existing data).')
    parser.add_argument('--port',type=int,default=5056)
    args=parser.parse_args()
    directory=ROOT/'.local'
    initialize(directory,args.seed)
    os.environ['MISB_DATA_DIR']=str(directory)
    os.environ['MISB_LOCAL_PREVIEW']='1'
    from app import app
    app.run(host='127.0.0.1',port=args.port,debug=False)
