import importlib.util
import io
import os
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile
from unittest.mock import patch
import openpyxl

ROOT=Path(__file__).resolve().parent.parent
sys.path.insert(0,str(ROOT/'src'))
sys.path.insert(0,str(ROOT/'scripts'))
from local import initialize
from app import app


class WebsiteTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp=tempfile.TemporaryDirectory()
        cls.root=Path(cls.temp.name)
        with zipfile.ZipFile(ROOT/'MISB_Fund_Tracker_simulation_copy_FIXED_v2.zip') as archive:
            archive.extractall(cls.root/'original')
        original=cls.root/'original/MISB_Fund_Tracker'
        spec=importlib.util.spec_from_file_location('original_app',original/'app.py')
        cls.original=importlib.util.module_from_spec(spec)
        spec.loader.exec_module(cls.original)
        cls.original.init_db()
        cls.original_payload=cls.original.build_payload()
        cls.seed=original/'data'

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def setUp(self):
        self.directory=tempfile.TemporaryDirectory()
        initialize(Path(self.directory.name),self.seed)
        self.env=patch.dict(os.environ,{'MISB_DATA_DIR':self.directory.name,'MISB_LOCAL_PREVIEW':'1'})
        self.env.start()
        self.client=app.test_client()

    def tearDown(self):
        self.env.stop()
        self.directory.cleanup()

    def test_original_dashboard_parity(self):
        response=self.client.get('/api/data')
        self.assertEqual(response.status_code,200)
        for key in ('summary','portfolio','issuers','profit_schedule','cash_projection','transactions'):
            # JSON round-trip normalizes date values exactly as both HTTP APIs do.
            with self.original.app.app_context():
                expected=self.original.app.json.loads(self.original.app.json.dumps(self.original_payload[key]))
            self.assertEqual(response.json[key],expected,key)

    def test_saved_changes_survive_new_client(self):
        self.assertEqual(self.client.post('/api/settings',json={'issuer_limit':1500000}).status_code,200)
        self.assertEqual(self.client.post('/api/monitoring/IIF-12345',json={'status':'Watch','owner':'Manager','notes':'Review'}).status_code,200)
        response=self.client.post('/api/cashflow-plan',json={'flow_type':'Allocation','flow_date':'2027-01-01','amount':1200})
        self.assertEqual(response.status_code,200)
        data=app.test_client().get('/api/data').json
        self.assertEqual(data['settings']['issuer_limit'],1500000)
        self.assertTrue(any(r['id']==response.json['id'] for r in data['cashflow_plan']))
        self.assertEqual(self.client.delete('/api/cashflow-plan/'+str(response.json['id'])).status_code,200)
        portfolio=data['portfolio']
        note=next(r for r in portfolio if any(not e['is_paid'] for e in r['payment_schedule']))
        event=next(e for e in note['payment_schedule'] if not e['is_paid'])
        mark={**event,'loan_code':note['loan_code'],'is_paid':True,'paid_date':'2026-08-24','paid_amount':123.45}
        self.assertEqual(self.client.post('/api/payment-mark',json=mark).status_code,200)
        refreshed=app.test_client().get('/api/data').json
        saved=next(r for r in refreshed['portfolio'] if r['loan_code']==note['loan_code'])
        payment=next(r for r in saved['payment_schedule'] if r['event_key']==event['event_key'])
        self.assertEqual(payment['paid_amount'],123.45)
        self.assertEqual(payment['paid_date'],'2026-08-24')

    def test_all_excel_exports(self):
        for endpoint in ('biweekly.xlsx','account-statement.xlsx?month=2026-08','updated-simulation.xlsx?as_of=2026-08-15'):
            response=self.client.get('/api/export/'+endpoint)
            self.assertEqual(response.status_code,200,endpoint)
            wb=openpyxl.load_workbook(io.BytesIO(response.data))
            self.assertTrue(wb.sheetnames)
            if endpoint=='biweekly.xlsx':
                cash=next(row[1].value for row in wb['Fund Summary'].iter_rows() if row[0].value=='Cash available')
                self.assertEqual(cash,self.original_payload['summary']['cash'])
            if endpoint.startswith('updated'):
                original=openpyxl.load_workbook(self.seed/'simulation.xlsx')
                self.assertEqual(wb.sheetnames,original.sheetnames)
                sheet=wb['Query result']; old=original['Query result']
                headers=[c.value for c in sheet[1]]
                col=headers.index('Remarks')+1
                self.assertEqual([sheet.cell(r,col).value for r in range(2,sheet.max_row+1)],
                                 [old.cell(r,col).value for r in range(2,old.max_row+1)])
                original.close()
            wb.close()

    def test_invalid_upload_does_not_replace_source(self):
        before=self.client.get('/api/data').json['summary']
        response=self.client.post('/api/upload/transactions',data={'file':(io.BytesIO(b'not excel'),'bad.xlsx')})
        self.assertEqual(response.status_code,400)
        self.assertEqual(self.client.get('/api/data').json['summary'],before)

    def test_valid_upload_preserves_backup(self):
        response=self.client.post('/api/upload/transactions',data={'file':(io.BytesIO((self.seed/'transactions.xlsx').read_bytes()),'transactions.xlsx')})
        self.assertEqual(response.status_code,200)
        self.assertTrue((Path(self.directory.name)/response.json['backup']).exists())
        self.assertEqual(self.client.get('/api/data').json['summary'],self.original_payload['summary'])

    def test_empty_workspace_accepts_first_import(self):
        with tempfile.TemporaryDirectory() as directory:
            initialize(Path(directory))
            with patch.dict(os.environ,{'MISB_DATA_DIR':directory}):
                response=self.client.get('/api/data')
                self.assertEqual(response.status_code,200)
                self.assertEqual(response.json['summary']['portfolio_count'],0)
                response=self.client.post('/api/upload/simulation',data={'file':(io.BytesIO((self.seed/'simulation.xlsx').read_bytes()),'simulation.xlsx'),'as_of':'2026-08-15'})
                self.assertEqual(response.status_code,200)

    def test_login_and_cross_origin_protection(self):
        with patch.dict(os.environ,{'MISB_LOCAL_PREVIEW':'0','APP_PASSWORD':'test-password-long','SESSION_SECRET':'x'*48}):
            self.assertEqual(self.client.get('/api/data').status_code,401)
            self.assertEqual(self.client.get('/').status_code,302)
            self.assertEqual(self.client.get('/login').status_code,200)
            self.assertEqual(self.client.post('/api/login',json={'password':'bad'}).status_code,401)
            self.assertEqual(self.client.post('/api/login',json={'password':'test-password-long'}).status_code,200)
            self.assertEqual(self.client.get('/api/data').status_code,200)
            self.assertEqual(self.client.post('/api/settings',json={'issuer_limit':100},headers={'Origin':'https://evil.example'}).status_code,403)
            self.assertEqual(self.client.post('/api/logout').status_code,200)
            self.assertEqual(self.client.get('/api/data').status_code,401)

    def test_invalid_json_and_static_privacy(self):
        self.assertEqual(self.client.post('/api/settings',json=[]).status_code,400)
        self.assertEqual(self.client.get('/data/transactions.xlsx').status_code,404)
        response=self.client.get('/')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.headers['Cache-Control'],'no-store')


if __name__=='__main__':
    unittest.main()
