import io
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
import openpyxl
from pypdf import PdfReader

ROOT=Path(__file__).resolve().parent.parent
sys.path.insert(0,str(ROOT/'src'));sys.path.insert(0,str(ROOT/'scripts'))
from statements import statement_data,statement_pdf
from app import app
from local import initialize


def workbook(rows):
    wb=openpyxl.Workbook();ws=wb.active
    ws.append(['Cofundr Admin Panel'])
    ws.append(['ID','Action','Sum Involved(MYR)','Previous Balance(MYR)','Current Balance(MYR)','Note #','Status','Date'])
    for row in reversed(rows): ws.append(row)
    output=io.BytesIO();wb.save(output);return output.getvalue()


ROWS=[
    [271150,'Deposit',1000,0,0,0,'SUCCESSFUL','21-Oct-2025 16:09:44'],
    [271151,'Deposit Approved',1000,0,1000,0,'SUCCESSFUL','21-Oct-2025 16:22:47'],
    [271152,'Investment Committed',200,1000,800,1996,'SUCCESSFUL','21-Oct-2025 17:10:14'],
    [271153,'Principal Payout',100,800,900,1996,'SUCCESSFUL','02-Sep-2026 12:00:00'],
    [271154,'Profit Payout',1.23,900,901.23,1996,'SUCCESSFUL','02-Sep-2026 12:00:00'],
    [271155,'Profit Payout',2.34,901.23,903.57,1996,'SUCCESSFUL','03-Sep-2026 12:00:00']]


class StatementTests(unittest.TestCase):
    def test_sst_starts_on_first_august_2026(self):
        rows=[
            [1,'Profit Payout',100,0,100,1996,'SUCCESSFUL','31-Jul-2026 23:59:59'],
            [2,'Profit Payout',100,100,200,1996,'SUCCESSFUL','01-Aug-2026 00:00:00']]
        result=statement_data(workbook(rows))
        self.assertEqual(result['summary']['sst'],'2.04')
        july=[r for r in result['rows'] if r['date'].startswith('31-Jul')]
        august=[r for r in result['rows'] if r['date'].startswith('01-Aug')]
        self.assertNotIn('SST',[r['description'] for r in july])
        self.assertIn('SST',[r['description'] for r in august])

    def test_inclusive_cutoff_and_approved_deposit(self):
        result=statement_data(workbook(ROWS),'2026-09-02')
        self.assertEqual(result['row_count'],7)
        self.assertEqual(result['rows'][0]['date'],'21-Oct-2025 16:22:47')
        self.assertEqual(result['summary'],dict(starting_balance='0.00',ending_balance='901.23',
            total_investment='200.00',principal_received='100.00',nett_returns='1.23',
            total_gross_returns='1.57',service_fee='0.31',sst='0.02'))
        self.assertEqual([row['description'] for row in result['rows'][-4:]],
                         ['Gross Profit','Service Charge','SST','Net Profit'])
        self.assertEqual(result['warnings'],[])
        self.assertEqual(statement_data(workbook(ROWS))['as_of'],'2026-09-03')

    def test_pending_deposits_excluded_and_invalid_dates_rejected(self):
        rows=ROWS+[[271156,'Deposit',500,903.57,903.57,0,'SUCCESSFUL','04-Sep-2026 12:00:00']]
        result=statement_data(workbook(rows))
        self.assertEqual(result['summary']['ending_balance'],'903.57')
        self.assertEqual(result['warnings'],[])
        with self.assertRaises(ValueError):statement_data(workbook(rows),'2026-02-30')
        with self.assertRaises(ValueError):statement_data(workbook(rows),'2025-01-01')

    def test_only_approved_withdrawal_uses_excel_timestamp(self):
        rows=ROWS+[[2711550,'Withdrawal',100,903.57,903.57,0,'SUCCESSFUL','04-Sep-2026 11:10:32'],[271156,'Withdrawal Approved',100,903.57,803.57,0,'SUCCESSFUL','04-Sep-2026 14:24:31'],[271157,'Withdrawal',1,803.57,802.57,0,'SUCCESSFUL','04-Sep-2026 14:24:31']]
        result=statement_data(workbook(rows))
        self.assertEqual(result['rows'][-2]['description'],'Withdrawal')
        self.assertEqual(result['rows'][-2]['date'],'04-Sep-2026 14:24:31')
        self.assertEqual(result['rows'][-1]['description'],'Withdrawal Fee')
        self.assertEqual(result['warnings'],[])
        self.assertEqual(result['rows'][-1]['date'],'04-Sep-2026 14:24:31')
        self.assertEqual(result['summary']['ending_balance'],'802.57')
        self.assertEqual(result['row_count'],13)

    def test_pdf_keeps_header_fonts_and_paginates(self):
        result=statement_data(workbook(ROWS),'2026-09-02')
        result['rows']=result['rows']*85
        pdf=PdfReader(statement_pdf(result))
        self.assertEqual(len(pdf.pages),11)
        first=pdf.pages[0].extract_text()
        self.assertIn('Crowd Sense Sdn Bhd',first)
        self.assertIn('Amanahraya Trustees Berhad',first)
        self.assertIn('ACCOUNT STATEMENT (YEAR TO DAY- 2 September 26)',first)
        self.assertIn('Total Gross Returns Received',first)
        self.assertIn('Service Fee',first)
        self.assertIn('901.23',first)
        self.assertIn('Calibri',str(pdf.pages[0]['/Resources']['/Font']['/T1']['/BaseFont']))
        self.assertEqual(tuple(pdf.pages[0].mediabox),(0,0,612,792))

    def test_upload_preview_download_and_version_guard(self):
        with tempfile.TemporaryDirectory() as directory:
            initialize(Path(directory))
            with patch.dict(os.environ,{'MISB_DATA_DIR':directory,'MISB_LOCAL_PREVIEW':'1'}):
                client=app.test_client()
                response=client.post('/api/account-statement',data={'file':(io.BytesIO(workbook(ROWS)),'ledger.xlsx'),'as_of':'2026-09-02'})
                self.assertEqual(response.status_code,200,response.data)
                self.assertEqual(response.json['row_count'],7)
                version=response.json['version']
                self.assertEqual(client.get('/api/account-statement-runs').json,[])
                response=client.get('/api/export/account-statement.pdf',query_string={'as_of':'2026-09-02','version':version})
                self.assertEqual(response.status_code,200,response.data[:100])
                self.assertEqual(response.mimetype,'application/pdf')
                self.assertTrue(response.data.startswith(b'%PDF-'))
                history=client.get('/api/account-statement-runs').json
                self.assertEqual(len(history),1)
                self.assertEqual(history[0]['as_of'],'2026-09-02')
                self.assertEqual(history[0]['source_object_key'],version)
                self.assertEqual(history[0]['row_count'],7)
                self.assertEqual(history[0]['closing_balance'],901.23)
                self.assertTrue(any(row['action']=='Generated account statement'
                                    for row in client.get('/api/audit-events').json))
                reproduced=client.get(f"/api/account-statement-runs/{history[0]['id']}/pdf")
                self.assertEqual(reproduced.status_code,200,reproduced.data[:100])
                self.assertEqual(reproduced.mimetype,'application/pdf')
                self.assertTrue(reproduced.data.startswith(b'%PDF-'))
                self.assertEqual(len(client.get('/api/account-statement-runs').json),1)
                self.assertEqual(client.get('/api/account-statement').json['version'],version)
                self.assertTrue(any(row['action']=='Recreated account statement'
                                    for row in client.get('/api/audit-events').json))
                self.assertEqual(client.get('/api/account-statement-runs/999999/pdf').status_code,404)
                self.assertEqual(client.get('/api/export/account-statement.pdf?version=stale').status_code,409)
                stale=client.post('/api/account-statement',data={
                    'file':(io.BytesIO(workbook(ROWS[:-1])),'older-ledger.xlsx'),'as_of':'2026-09-02'})
                self.assertEqual(stale.status_code,409,stale.data)
                self.assertTrue(stale.json['comparison']['requires_acknowledgement'])
                self.assertEqual(client.get('/api/account-statement').json['version'],version)
                response=client.post('/api/account-statement',data={'file':(io.BytesIO(b'bad file'),'bad.xlsx')})
                self.assertEqual(response.status_code,400)
                self.assertEqual(client.get('/api/account-statement').json['version'],version)
