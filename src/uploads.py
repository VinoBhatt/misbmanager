"""Bound spreadsheet expansion before handing XML to openpyxl."""
import io
import zipfile
from datetime import date, datetime
import openpyxl


def validate_workbook(kind, data):
    if len(data) > 10 * 1024 * 1024:
        raise ValueError('Workbook exceeds 10 MB')
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        entries = archive.infolist()
        if len(entries) > 2000 or sum(e.file_size for e in entries) > 40 * 1024 * 1024:
            raise ValueError('Expanded workbook is too large')
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        if kind == 'transactions':
            sheet = wb.worksheets[0]
            headers = next(sheet.iter_rows(min_row=2, max_row=2, values_only=True))
            required = {'ID','Action','Sum Involved(MYR)','Previous Balance(MYR)',
                        'Current Balance(MYR)','Note #','Status','Date'}
        elif kind == 'simulation':
            if 'Query result' not in wb.sheetnames:
                raise ValueError('Missing Query result sheet')
            sheet = wb['Query result']
            headers = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
            required = {'Loan Code','Issuer Name','Investment Amount','Loan Status',
                        'Paid Principal','Unpaid Principal','Paid Profit','Unpaid Profit'}
        else:
            if not {'Sheet1','Sheet2'}.issubset(wb.sheetnames):
                raise ValueError('Missing MIDAS sheets')
            return
        if not required.issubset(set(headers)):
            raise ValueError('Required columns are missing')
        if kind == 'simulation':
            legacy={'Gross Profit','Late Profit','Service Fee ','Net Profit'}
            expanded={'Gross Profit Earned','Total Gross Profit','Late Payment Charges','Service Fee ','SST','Net Profit','Installment'}
            if not (legacy.issubset(set(headers)) or expanded.issubset(set(headers))):
                raise ValueError('The simulation profit columns are not a supported format')
        if (sheet.max_row or 0) > 50000 or (sheet.max_column or 0) > 500:
            raise ValueError('Workbook dimensions exceed supported limits')
    finally:
        wb.close()


def inspect_workbook(kind, data):
    """Return a concise, non-mutating quality report for an import candidate."""
    validate_workbook(kind,data)
    wb=openpyxl.load_workbook(io.BytesIO(data),read_only=True,data_only=True)
    issues=[];stats={'bytes':len(data),'sheets':len(wb.sheetnames)}
    def issue(level,message,count=1):
        issues.append({'level':level,'message':message,'count':count})
    def parsed_date(value):
        if isinstance(value,(datetime,date)): return value
        for fmt in ('%d-%b-%Y %H:%M:%S','%d/%m/%Y','%Y-%m-%d %H:%M:%S','%Y-%m-%d'):
            try: return datetime.strptime(str(value or '').strip(),fmt)
            except ValueError: pass
        return None
    try:
        if kind=='transactions':
            ws=wb.worksheets[0];headers=next(ws.iter_rows(min_row=2,max_row=2,values_only=True))
            rows=[]
            for values in ws.iter_rows(min_row=3,values_only=True):
                if any(value is not None for value in values): rows.append(dict(zip(headers,values)))
            successful=[row for row in rows if str(row.get('Status','')).upper()=='SUCCESSFUL']
            stats.update(rows=len(rows),successful=len(successful),pending=len(rows)-len(successful))
            ids=[row.get('ID') for row in rows if row.get('ID') not in (None,'')]
            duplicates=len(ids)-len(set(ids))
            if duplicates: issue('error','Duplicate transaction IDs',duplicates)
            invalid_dates=sum(parsed_date(row.get('Date')) is None for row in successful)
            if invalid_dates: issue('error','Successful transactions with invalid dates',invalid_dates)
            invalid_amounts=sum(not isinstance(row.get('Sum Involved(MYR)'),(int,float)) for row in successful)
            if invalid_amounts: issue('error','Successful transactions with invalid amounts',invalid_amounts)
            note_actions={'Investment Committed','Principal Payout','Profit Payout'}
            missing_notes=sum(row.get('Action') in note_actions and row.get('Note #') in (None,'',0,'0') for row in successful)
            if missing_notes: issue('warning','Note transactions without a note number',missing_notes)
            approved=sum(row.get('Action') in ('Deposit Approved','Withdrawal Approved') for row in successful)
            stats['approved_cash_movements']=approved
        elif kind=='simulation':
            ws=wb['Query result'];headers=next(ws.iter_rows(min_row=1,max_row=1,values_only=True))
            rows=[dict(zip(headers,values)) for values in ws.iter_rows(min_row=2,values_only=True)
                  if values and values[0] is not None and values[1]]
            codes=[str(row.get('Loan Code') or '').strip().upper() for row in rows]
            duplicates=len(codes)-len(set(codes))
            missing_issuers=sum(not str(row.get('Issuer Name') or '').strip() for row in rows)
            invalid_amounts=sum(not isinstance(row.get('Investment Amount'),(int,float)) or row.get('Investment Amount',0)<0 for row in rows)
            stats.update(rows=len(rows),format=('Expanded profit' if 'Total Gross Profit' in headers else 'Original'))
            if duplicates: issue('error','Duplicate loan codes',duplicates)
            if missing_issuers: issue('warning','Notes without an issuer name',missing_issuers)
            if invalid_amounts: issue('error','Notes with invalid investment amounts',invalid_amounts)
        else:
            stats.update(rows=(wb['Sheet2'].max_row or 0),format='MIDAS projections')
        if not issues: issue('ok','No blocking data-quality issues found',0)
        return {'kind':kind,'stats':stats,'issues':issues,
                'can_import':not any(item['level']=='error' for item in issues)}
    finally:
        wb.close()
