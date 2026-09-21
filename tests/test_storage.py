from contextlib import closing
import os
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT/'src'))
sys.path.insert(0, str(ROOT/'scripts'))
from app import app
from local import initialize
from storage import CHUNK_BYTES, atomic_write, read_source, save_source, workbook_statements


class WorkbookStorageTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp.name)
        initialize(self.directory)
        self.env = patch.dict(os.environ, {'MISB_DATA_DIR':self.temp.name})
        self.env.start()

    def tearDown(self):
        self.env.stop()
        self.temp.cleanup()

    def test_multiple_chunks_roundtrip_and_backup(self):
        data = os.urandom(CHUNK_BYTES*2+71)
        with app.test_request_context():
            self.assertIsNone(save_source('simulation',data,'2026-09-21'))
        with app.test_request_context():
            self.assertEqual(read_source('simulation').read(),data)
            previous = save_source('simulation',b'replacement','2026-09-22')
        with closing(sqlite3.connect(self.directory/'misb_tracker.db')) as con, con:
            self.assertEqual(con.execute('SELECT COUNT(*) FROM workbook_chunks WHERE object_key=?',(previous,)).fetchone()[0],3)
        with app.test_request_context():
            self.assertEqual(read_source('simulation').read(),b'replacement')

    def test_failed_transaction_leaves_no_partial_version(self):
        with app.test_request_context():
            save_source('transactions',b'original','2026-09-21')
            statements = workbook_statements('broken',b'new bytes')
            statements.append(('INSERT INTO table_that_does_not_exist VALUES(?)',(1,)))
            with self.assertRaises(sqlite3.OperationalError):
                atomic_write(statements)
        with closing(sqlite3.connect(self.directory/'misb_tracker.db')) as con, con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM workbook_versions WHERE object_key='broken'").fetchone()[0],0)
        with app.test_request_context():
            self.assertEqual(read_source('transactions').read(),b'original')

    def test_missing_or_corrupt_chunk_rejected(self):
        with app.test_request_context():
            save_source('simulation',b'original','2026-09-21')
        with closing(sqlite3.connect(self.directory/'misb_tracker.db')) as con, con:
            con.execute("UPDATE workbook_chunks SET content='dGFtcGVyZWQ='")
        with app.test_request_context():
            with self.assertRaisesRegex(ValueError,'integrity'):
                read_source('simulation')

    def test_existing_local_file_migrated(self):
        target = self.directory/'sources/old.xlsx'
        target.parent.mkdir();target.write_bytes(b'old workbook')
        with closing(sqlite3.connect(self.directory/'misb_tracker.db')) as con, con:
            con.execute("INSERT INTO sources(kind,object_key) VALUES('simulation','sources/old.xlsx')")
        initialize(self.directory)
        initialize(self.directory)
        with app.test_request_context():
            self.assertEqual(read_source('simulation').read(),b'old workbook')
