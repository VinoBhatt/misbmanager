import io
import csv
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
    def test_csv_download_matches_prepared_monthly_statement(self):
        with tempfile.TemporaryDirectory() as directory:
            initialize(Path(directory))
            with patch.dict(os.environ,{'MISB_DATA_DIR':directory,'MISB_LOCAL_PREVIEW':'1'}):
                client=app.test_client()
                uploaded=client.post('/api/account-statement',data={
                    'file':(io.BytesIO(workbook(ROWS)),'ledger.xlsx')})
                params={'mode':'monthly','start_date':'2026-09-03',
                        'as_of':'2026-09-30','version':uploaded.json['version']}
                preview=client.get('/api/account-statement',query_string=params).json
                response=client.get('/api/export/account-statement.csv',query_string=params)
                self.assertEqual(response.status_code,200)
                self.assertEqual(response.mimetype,'text/csv')
                self.assertIn('MISB_Account_Statement_2026-09-03_to_2026-09-30.csv',
                              response.headers['Content-Disposition'])
                rows=list(csv.reader(io.StringIO(response.data.decode('utf-8-sig'))))
                self.assertEqual(len(rows),preview['row_count']+1)
                self.assertEqual(rows[1:],[[row[key] for key in
                    ('date','description','note','previous','amount','current')]
                    for row in preview['rows']])
                self.assertEqual(rows[-1][-1],'903.57')
                self.assertTrue(any(row[4].startswith('-') for row in rows[1:]))
                self.assertEqual(client.get('/api/account-statement-runs').json,[])
                self.assertEqual(client.get('/api/export/account-statement.csv?version=stale').status_code,409)
                self.assertEqual(client.get('/api/export/account-statement.csv?as_of=invalid').status_code,400)

    def test_monthly_carries_september_second_balance(self):
        result=statement_data(workbook(ROWS),'2026-09-30','2026-09-03')
        self.assertEqual(result['summary']['starting_balance'],'901.23')
        self.assertEqual(result['summary']['ending_balance'],'903.57')
        self.assertEqual(result['summary']['principal_received'],'0.00')
        self.assertEqual(result['row_count'],4)
        self.assertTrue(all(row['date'].startswith('03-Sep') for row in result['rows']))
        with self.assertRaisesRegex(ValueError,'does not match'):
            statement_data(workbook(ROWS),'2026-09-30','2026-09-03','999.00')

    def test_monthly_history_reproduces_and_verifies_period(self):
        with tempfile.TemporaryDirectory() as directory:
            initialize(Path(directory))
            with patch.dict(os.environ,{'MISB_DATA_DIR':directory,'MISB_LOCAL_PREVIEW':'1'}):
                client=app.test_client()
                response=client.post('/api/account-statement',data={
                    'file':(io.BytesIO(workbook(ROWS)),'ledger.xlsx'),'as_of':'2026-09-02'})
                version=response.json['version']
                client.get('/api/export/account-statement.pdf',query_string={'as_of':'2026-09-02','version':version})
                params={'mode':'monthly','as_of':'2026-09-30','version':version}
                preview=client.get('/api/account-statement',query_string=params)
                self.assertEqual(preview.status_code,200,preview.data)
                self.assertEqual(preview.json['start_date'],'2026-09-03')
                self.assertEqual(preview.json['summary']['starting_balance'],'901.23')
                params['start_date']=preview.json['start_date']
                exported=client.get('/api/export/account-statement.pdf',query_string=params)
                self.assertEqual(exported.status_code,200,exported.data[:100])
                run_id=exported.headers['X-Statement-History-ID']
                self.assertIn('2026-09-03 TO 2026-09-30',PdfReader(io.BytesIO(exported.data)).pages[0].extract_text())
                self.assertTrue(client.post(f'/api/account-statement-runs/{run_id}/verify').json['matches'])
                self.assertEqual(client.get(f'/api/account-statement-runs/{run_id}/pdf').status_code,200)
                self.assertEqual(client.get('/api/export/account-statement.pdf',query_string=params).headers['X-Statement-History-New'],'0')

    def test_statement_upload_can_overwrite_balance_gaps(self):
        rows=[list(row) for row in ROWS]
        rows[-1][3]=902
        rows[-1][4]=904.34
        payload=workbook(rows)
        with tempfile.TemporaryDirectory() as directory:
            initialize(Path(directory))
            with patch.dict(os.environ,{'MISB_DATA_DIR':directory,'MISB_LOCAL_PREVIEW':'1'}):
                client=app.test_client()
                rejected=client.post('/api/account-statement',data={
                    'file':(io.BytesIO(payload),'latest.xlsx')})
                self.assertEqual(rejected.status_code,422)
                imported=client.post('/api/account-statement',data={
                    'file':(io.BytesIO(payload),'latest.xlsx'),'overwrite':'1','allow_regression':'1'})
                self.assertEqual(imported.status_code,200,imported.data)
                self.assertEqual(imported.json['summary']['ending_balance'],'904.34')
                self.assertTrue(any('Imported ledger warning' in warning for warning in imported.json['warnings']))

    def test_request_export_explains_wrong_workbook(self):
        from uploads import inspect_workbook
        wb=openpyxl.Workbook();ws=wb.active
        ws.append(['Cofundr Admin Panel'])
        ws.append(['#','Note ID','Name','User Type','Bank','Account No','Amount(MYR)',
                   'Transaction No','Transaction Proof','Date','Pay Via','Action'])
        output=io.BytesIO();wb.save(output);wb.close()
        with self.assertRaisesRegex(ValueError,'request export, not an account transaction log'):
            inspect_workbook('transactions',output.getvalue())

    def test_user_example_sst_is_eight_percent_of_platform_fee(self):
        result=statement_data(workbook([
            [1,'Profit Payout',78.40,0,78.40,1996,'SUCCESSFUL','02-Sep-2026 12:00:00']]))
        self.assertEqual(result['summary']['total_gross_returns'],'100.00')
        self.assertEqual(result['summary']['service_fee'],'20.00')
        self.assertEqual(result['summary']['sst'],'1.60')
        self.assertEqual(result['summary']['nett_returns'],'78.40')
        self.assertEqual([row['amount'] for row in result['rows']],['100.00','-20.00','-1.60','78.40'])

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
                ledger_bytes=workbook(ROWS)
                response=client.post('/api/account-statement',data={'file':(io.BytesIO(ledger_bytes),'ledger.xlsx'),'as_of':'2026-09-02'})
                self.assertEqual(response.status_code,200,response.data)
                self.assertEqual(response.json['row_count'],7)
                version=response.json['version']
                self.assertEqual(client.get('/api/account-statement-runs').json,[])
                response=client.get('/api/export/account-statement.pdf',query_string={'as_of':'2026-09-02','version':version})
                self.assertEqual(response.status_code,200,response.data[:100])
                self.assertEqual(response.mimetype,'application/pdf')
                self.assertTrue(response.data.startswith(b'%PDF-'))
                self.assertEqual(response.headers['X-Statement-History-New'],'1')
                history=client.get('/api/account-statement-runs').json
                self.assertEqual(len(history),1)
                self.assertEqual(history[0]['as_of'],'2026-09-02')
                self.assertEqual(history[0]['source_object_key'],version)
                self.assertEqual(history[0]['row_count'],7)
                self.assertEqual(history[0]['closing_balance'],901.23)
                self.assertEqual(history[0]['source_filename'],'ledger.xlsx')
                self.assertEqual(history[0]['source_available'],1)
                self.assertEqual(len(history[0]['source_sha256']),64)
                self.assertGreater(history[0]['source_byte_size'],0)
                register=client.get('/api/account-statement-runs.csv')
                self.assertEqual(register.status_code,200)
                self.assertEqual(register.mimetype,'text/csv')
                register_rows=list(csv.reader(io.StringIO(register.data.decode('utf-8-sig'))))
                self.assertEqual(len(register_rows),2)
                self.assertEqual(register_rows[0][0:3],['Generated at','Statement through','PDF filename'])
                self.assertEqual(register_rows[1][1],'2026-09-02')
                self.assertEqual(register_rows[1][8],'ledger.xlsx')
                self.assertEqual(register_rows[1][9],history[0]['source_sha256'])
                self.assertEqual(register_rows[1][11],'Yes')
                self.assertTrue(any(row['action']=='Generated account statement'
                                    for row in client.get('/api/audit-events').json))
                verified=client.post(f"/api/account-statement-runs/{history[0]['id']}/verify")
                self.assertEqual(verified.status_code,200,verified.data)
                self.assertTrue(verified.json['matches'])
                self.assertTrue(all(check['match'] for check in verified.json['checks'].values()))
                self.assertEqual(verified.json['checks']['Closing balance']['calculated'],'901.23')
                self.assertEqual(verified.json['source_sha256'],history[0]['source_sha256'])
                self.assertTrue(verified.json['verified_at'])
                verified_history=client.get('/api/account-statement-runs').json
                self.assertEqual(verified_history[0]['last_verification_matches'],1)
                self.assertEqual(verified_history[0]['last_verification_mismatches'],[])
                self.assertEqual(verified_history[0]['last_verified_at'],verified.json['verified_at'])
                verified_register=list(csv.reader(io.StringIO(
                    client.get('/api/account-statement-runs.csv').data.decode('utf-8-sig'))))
                self.assertEqual(verified_register[1][12],verified.json['verified_at'])
                self.assertEqual(verified_register[1][13],'Matches')
                self.assertEqual(verified_register[1][14],'')
                self.assertTrue(any(row['action']=='Verified account statement'
                                    for row in client.get('/api/audit-events').json))
                self.assertEqual(client.post('/api/account-statement-runs/999999/verify').status_code,404)
                repeated=client.get('/api/export/account-statement.pdf',
                                    query_string={'as_of':'2026-09-02','version':version})
                self.assertEqual(repeated.status_code,200)
                self.assertEqual(repeated.headers['X-Statement-History-New'],'0')
                self.assertEqual(repeated.headers['X-Statement-History-ID'],str(history[0]['id']))
                self.assertEqual(len(client.get('/api/account-statement-runs').json),1)
                self.assertTrue(any(row['action']=='Downloaded existing account statement'
                                    for row in client.get('/api/audit-events').json))
                reproduced=client.get(f"/api/account-statement-runs/{history[0]['id']}/pdf")
                self.assertEqual(reproduced.status_code,200,reproduced.data[:100])
                self.assertEqual(reproduced.mimetype,'application/pdf')
                self.assertTrue(reproduced.data.startswith(b'%PDF-'))
                self.assertEqual(reproduced.headers['X-Statement-Source-SHA256'],history[0]['source_sha256'])
                source=client.get('/api/source-versions/transactions/download',
                                  query_string={'object_key':history[0]['source_object_key']})
                self.assertEqual(source.status_code,200)
                self.assertEqual(source.data,ledger_bytes)
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
