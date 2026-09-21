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
    con.executescript((ROOT/'migrations/0001_initial.sql').read_text())
    if seed:
        import shutil
        import json
        mapping={'transactions':'transactions.xlsx','simulation':'simulation.xlsx','projections':'midas_projections.xlsx'}
        for kind,filename in mapping.items():
            if con.execute('SELECT 1 FROM sources WHERE kind=?',(kind,)).fetchone():
                continue
            source=seed/filename
            if source.exists():
                key='sources/initial/'+filename
                target=directory/key
                target.parent.mkdir(parents=True,exist_ok=True)
                shutil.copyfile(source,target)
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
