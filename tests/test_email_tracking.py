import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from flask import g
import openpyxl
import io

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT/'src'), str(ROOT/'scripts')]
from app import app, build_updated_simulation_workbook, simulation_update_snapshot
from local import initialize
from storage import save_source


class EmailTrackingTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        initialize(Path(self.directory.name))
        self.env = patch.dict(os.environ, {'MISB_DATA_DIR': self.directory.name, 'MISB_LOCAL_PREVIEW': '1'})
        self.env.start()
        self.client = app.test_client()
        self.transactions = [self.transaction(1)]
        self.headers = ['No', 'Loan Code', 'Issuer Name', 'Email on Allocation', 'Email on Disbursement', 'Disbursal Date']
        self.values = [1, 'IIF-1234', 'Example issuer', '2026-01-01', 'Nil', '2026-01-03']
        self.tx_patch = patch('app.load_transactions', side_effect=lambda: self.transactions)
        self.sim_patch = patch('app.load_simulation', side_effect=self.simulation)
        self.tx_patch.start(); self.sim_patch.start()

    def tearDown(self):
        self.sim_patch.stop(); self.tx_patch.stop(); self.env.stop(); self.directory.cleanup()

    def transaction(self, identifier, code='IIF-1234', status='SUCCESSFUL', day='2026-01-01'):
        return {'id': identifier, 'note': code, 'status': status, 'date': day,
                'amount': 100, 'action': 'Investment Committed', 'transaction_no': str(identifier)}

    def simulation(self):
        g.simulation_display = {'headers': self.headers, 'rows': [self.values]}
        return [{'loan_code': 'IIF-1234', 'investment_amount': 100, 'loan_status': 'Active'}], []

    def register(self):
        response = self.client.get('/api/allocation-email-tracking')
        self.assertEqual(response.status_code, 200)
        return response.json

    def save(self, **overrides):
        row = next(r for r in self.register()['rows'] if r['loan_code'] == 'IIF-1234')
        return self.client.put('/api/allocation-email-tracking/IIF-1234', json={
            'allocation_email_date': '2026-01-01', 'disbursement_email_date': '2026-01-03',
            'ledger_signature': row['ledger_signature'], **overrides})

    def test_register_uses_successful_allocations_and_keeps_missing_notes(self):
        self.transactions += [self.transaction(2, 'IIF-5678'), self.transaction(3, status='FAILED'), self.transaction(4, '')]
        data = self.register()
        self.assertEqual(data['summary']['total'], 2)
        self.assertEqual(len(data['unmatched']), 1)
        row = next(r for r in data['rows'] if r['loan_code'] == 'IIF-1234')
        self.assertEqual(row['allocation_email_date'], '2026-01-01')
        self.assertEqual(row['disbursement_email_date'], '')
        self.assertEqual(row['amount'], 100)
        missing = next(r for r in data['rows'] if r['loan_code'] == 'IIF-5678')
        self.assertEqual(missing['source'], 'Not in simulation')
        self.assertIn('Disbursement date missing', missing['pending'])

    def test_saved_dates_persist_clear_and_reopen_on_same_day_top_up(self):
        self.assertEqual(self.save().status_code, 200)
        self.assertTrue(app.test_client().get('/api/allocation-email-tracking').json['rows'][0]['emails_complete'])
        self.transactions.append(self.transaction(2))
        row = self.register()['rows'][0]
        self.assertFalse(row['emails_complete'])
        self.assertIn('Review allocation coverage', row['pending'])
        self.assertEqual(self.save().status_code, 200)
        self.assertTrue(self.register()['rows'][0]['emails_complete'])
        self.assertEqual(self.save(disbursement_email_date='').status_code, 200)
        self.assertFalse(self.register()['rows'][0]['emails_complete'])
        self.assertEqual(self.register()['rows'][0]['disbursement_email_date'], '')

    def test_rejects_stale_confirmation_invalid_dates_future_dates_and_unknown_notes(self):
        self.assertEqual(self.save(ledger_signature='outdated').status_code, 409)
        self.assertEqual(self.save(disbursement_email_date='not a date').status_code, 400)
        self.assertEqual(self.save(disbursement_email_date='2999-01-01').status_code, 400)
        self.assertEqual(self.client.put('/api/allocation-email-tracking/IIF-9999', json={}).status_code, 404)
        self.assertEqual(self.register()['rows'][0]['source'], 'Imported simulation')

    def test_imported_email_dates_before_later_allocation_require_review(self):
        self.values[4] = '2026-01-03'
        self.transactions.append(self.transaction(2, day='2026-01-05'))
        row = self.register()['rows'][0]
        self.assertFalse(row['emails_complete'])
        self.assertIn('Review allocation coverage', row['pending'])

    def test_preview_and_export_use_confirmed_email_dates(self):
        self.assertEqual(self.save().status_code, 200)
        workbook = openpyxl.Workbook(); sheet = workbook.active; sheet.title = 'Query result'
        sheet.append(self.headers); sheet.append(self.values)
        content = io.BytesIO(); workbook.save(content)
        with app.test_request_context():
            save_source('simulation', content.getvalue(), '2026-01-03')
            snapshot = simulation_update_snapshot()
            self.assertEqual(snapshot['rows'][0][4], '2026-01-03')
            exported, _ = build_updated_simulation_workbook(snap=snapshot)
            self.assertEqual(exported['Query result']['E2'].value, '2026-01-03')
            self.assertEqual(exported['Query result']['F2'].value, '2026-01-03')
        self.assertEqual(self.save(disbursement_email_date='').status_code, 200)
        with app.test_request_context():
            snapshot = simulation_update_snapshot()
            exported, _ = build_updated_simulation_workbook(snap=snapshot)
            self.assertEqual(exported['Query result']['E2'].value, 'Nil')


if __name__ == '__main__':
    unittest.main()
