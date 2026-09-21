"""Export original user records for a one-time import into a new D1 database."""
import argparse
from pathlib import Path
import sqlite3

parser=argparse.ArgumentParser()
parser.add_argument('database',type=Path)
args=parser.parse_args()
if not args.database.is_file():
    parser.error('Database does not exist')
con=sqlite3.connect(args.database.resolve().as_uri()+'?mode=ro',uri=True)
allowed=('monitoring','settings','payment_marks','cashflow_plan')
lines=[]
for line in con.iterdump():
    if any(line.startswith('INSERT INTO "'+table+'"') for table in allowed):
        lines.append(line.replace('INSERT INTO','INSERT OR REPLACE INTO',1))
con.close()
output=Path('.local/legacy-records.sql')
output.parent.mkdir(exist_ok=True)
output.write_text('\n'.join(lines)+'\n',encoding='utf-8')
print(f'Exported {len(lines)} records to {output}. Review before importing.')
