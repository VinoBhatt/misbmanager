"""MISB PDF statements: exact decimal ledger amounts and reference artwork."""
import io
from zipfile import BadZipFile
from datetime import datetime, date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import openpyxl

CENT = Decimal('0.01')
MONTHS = ('January','February','March','April','May','June','July','August','September','October','November','December')


def money(value):
    try:
        amount = Decimal(str(value))
        if not amount.is_finite():
            raise ValueError('Amounts must be finite')
        return amount.quantize(CENT, rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError):
        raise ValueError('A transaction has an invalid amount') from None


def timestamp(value):
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    for fmt in ('%d-%b-%Y %H:%M:%S','%Y-%m-%d %H:%M:%S','%Y-%m-%d','%d/%m/%Y'):
        try:
            return datetime.strptime(str(value).strip(),fmt)
        except ValueError:
            pass
    raise ValueError('A successful transaction has an invalid date')


def statement_data(data, as_of=None):
    wb = openpyxl.load_workbook(io.BytesIO(data),data_only=True,read_only=True)
    transactions = []
    try:
        ws = wb.worksheets[0]
        headers = next(ws.iter_rows(min_row=2,max_row=2,values_only=True))
        for index, values in enumerate(ws.iter_rows(min_row=3,values_only=True)):
            row = dict(zip(headers,values))
            if str(row.get('Status','')).upper() != 'SUCCESSFUL':
                continue
            dt = timestamp(row.get('Date'))
            note = row.get('Note #')
            note = '-' if note in (None,'',0,'0') else str(note).removeprefix('IIF-')
            if note.endswith('.0'): note=note[:-2]
            transactions.append({'timestamp':dt,'order':int(row.get('ID') or index),
                'description':str(row.get('Action') or '').strip(),'note':note,
                'previous':money(row.get('Previous Balance(MYR)')),
                'amount':money(row.get('Sum Involved(MYR)')),
                'current':money(row.get('Current Balance(MYR)'))})
            if len(transactions)>10000:
                raise ValueError('Statements support up to 10,000 transactions per workbook')
    finally:
        wb.close()
    if not transactions:
        raise ValueError('The workbook has no successful transactions')
    transactions.sort(key=lambda r:(r['timestamp'],r['order']))
    try:
        cutoff = date.fromisoformat(as_of) if as_of else transactions[-1]['timestamp'].date()
    except (ValueError,TypeError):
        raise ValueError('Cut-off date must be YYYY-MM-DD') from None
    selected = [r for r in transactions if r['timestamp'].date() <= cutoff]
    rows, warnings = [], []
    for r in selected:
        r = dict(r)
        action = r['description']
        if action == 'Deposit':
            continue
        if action == 'Withdrawal':
            # The export records the RM1 fee as a separate Withdrawal row.
            # Match its balance movement to the approved withdrawal, not a request.
            fee = (r['amount'] == Decimal('1.00')
                   and r['previous'] - r['current'] == r['amount']
                   and any(a['description'] == 'Withdrawal Approved'
                           and a['timestamp'] == r['timestamp']
                           and a['current'] == r['previous'] for a in selected))
            if not fee:
                continue
            r['description'] = 'Withdrawal Fee'
        if action in ('Deposit Approved', 'Withdrawal Approved'):
            r['description'] = action.removesuffix(' Approved')
            if r['note'] == '-':
                r['note'] = '0'
        rows.append(r)
    if not rows:
        raise ValueError('No completed cash transactions exist on or before this cut-off date')
    breaks = sum(a['current'] != b['previous'] for a,b in zip(rows,rows[1:]))
    if breaks:
        warnings.append(f'{breaks} balance discontinuity/discontinuities found. The PDF preserves the original ledger balances; check that the log is complete.')
    totals = {
        'starting_balance':rows[0]['previous'],'ending_balance':rows[-1]['current'],
        'total_investment':sum((r['amount'] for r in rows if r['description']=='Investment Committed'),Decimal(0)),
        'principal_received':sum((r['amount'] for r in rows if r['description']=='Principal Payout'),Decimal(0)),
        'nett_returns':sum((r['amount'] for r in rows if r['description']=='Profit Payout'),Decimal(0))}
    return {'as_of':cutoff.isoformat(),'latest_date':transactions[-1]['timestamp'].date().isoformat(),
        'investor_id':'5490','investor_name':'Amanahraya Trustees Berhad',
        'summary':{k:format(v,'.2f') for k,v in totals.items()},
        'rows':[{'date':r['timestamp'].strftime('%d-%b-%Y %H:%M:%S'),'description':r['description'],
                 'note':r['note'],**{k:format(r[k],'.2f') for k in ('previous','amount','current')}} for r in rows],
        'row_count':len(rows),'page_count':1+max(0,(len(rows)-29+54)//55),'warnings':warnings}


def statement_pdf(statement):
    # Lazy imports keep the Worker startup path small.
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import NameObject, DecodedStreamObject
    from statement_template import TEMPLATE
    reader=PdfReader(io.BytesIO(TEMPLATE)); writer=PdfWriter()
    first=writer.add_page(reader.pages[0])
    fonts=first['/Resources']['/Font']

    def width(text,font,size):
        f=fonts[font].get_object(); widths=f['/Widths']; start=int(f['/FirstChar'])
        encoded=text.encode('cp1252')
        return sum(float(widths[c-start]) if 0<=c-start<len(widths) else 0 for c in encoded)*size/1000

    def text(text,x,y,size=6.96,font='/F1',align='left'):
        w=width(text,font,size)
        if align=='right': x-=w
        if align=='center': x-=w/2
        literal=text.encode('cp1252').hex()
        return f'BT 0 g {font} {size} Tf 1 0 0 1 {x:.4f} {792-y:.4f} Tm <{literal}> Tj ET\n'

    cutoff=date.fromisoformat(statement['as_of'])
    title=f'ACCOUNT STATEMENT (YEAR TO DAY- {cutoff.day} {MONTHS[cutoff.month-1]} {cutoff:%y})'
    first_ops=text(title,302.9,197.69,11.4,'/F2','center')
    for key,y in zip(('starting_balance','ending_balance','total_investment','principal_received','nett_returns'),(272.57,284.81,297.05,309.29,321.55)):
        first_ops+=text(format(Decimal(statement['summary'][key]),',.2f'),303.31,y,7.56,align='right')
    all_rows=statement['rows']; groups=[all_rows[:29]]+[all_rows[i:i+55] for i in range(29,len(all_rows),55)]
    columns=(50.964,154.31,306.49,352.93,419.8,486.64,553.5)
    for page_index,rows in enumerate(groups):
        page=first if page_index==0 else writer.add_blank_page(width=612,height=792)
        if page_index: page[NameObject('/Resources')]=first['/Resources']
        ops=first_ops if page_index==0 else ''
        top=380.17 if page_index==0 else 54.56
        bottom=top+len(rows)*12.24
        ops+='0 G 0.6 w\n'
        for x in columns:
            ops+=f'{x} {792-top:.4f} m {x} {792-bottom:.4f} l S\n'
        for i in range(len(rows)+1):
            y=792-top-i*12.24
            ops+=f'{columns[0]} {y:.4f} m {columns[-1]} {y:.4f} l S\n'
        for i,row in enumerate(rows):
            y=(388.51 if page_index==0 else 62.90)+12.24*i
            for value,left,right in zip((row['date'],row['description'],row['note']),columns[:3],columns[1:4]):
                size=min(6.96,6.96*(right-left-3)/max(1,width(value,'/F1',6.96)))
                ops+=text(value,(left+right)/2,y,size,align='center')
            for key,right in zip(('previous','amount','current'),(417.55,484.39,551.25)):
                ops+=text(format(Decimal(row[key]),',.2f'),right,y,align='right')
        old=page.get_contents()
        stream=DecodedStreamObject()
        stream.set_data((b'q\n'+old.get_data()+b'\nQ\n' if old is not None else b'')+ops.encode('ascii'))
        page[NameObject('/Contents')]=stream
        page.compress_content_streams()
    writer.add_metadata({'/Title':f'MISB Account Statement through {statement["as_of"]}', '/Author':'Crowd Sense Sdn Bhd'})
    output=io.BytesIO();writer.write(output);output.seek(0)
    return output


def register_statement_routes(app):
    from flask import request, jsonify, send_file
    from storage import read_source, source_rows, save_source, source_exists
    from uploads import validate_workbook

    def build():
        if not source_exists('transactions'):
            raise ValueError('Upload a transaction log first')
        expected=request.args.get('version')
        if expected and expected!=source_rows()['transactions']['object_key']:
            return None
        result=statement_data(read_source('transactions').read(),request.args.get('as_of'))
        result['version']=source_rows()['transactions']['object_key']
        return result

    @app.route('/api/account-statement',methods=['GET','POST'])
    def preview_statement():
        try:
            if request.method=='POST':
                file=request.files.get('file')
                if not file or not file.filename.lower().endswith('.xlsx'):
                    raise ValueError('Choose a transaction log in .xlsx format')
                data=file.read();validate_workbook('transactions',data)
                result=statement_data(data,request.form.get('as_of'))
                save_source('transactions',data,result['latest_date'])
                result['version']=source_rows()['transactions']['object_key']
            else:
                result=build()
            if result is None: return jsonify(error='The ledger has changed. Refresh the statement preview.'),409
            return jsonify(result)
        except (BadZipFile,openpyxl.utils.exceptions.InvalidFileException):
            return jsonify(error='This file is not a readable Excel workbook.'),400
        except (ValueError,KeyError,StopIteration) as error:
            return jsonify(error=str(error)),400

    @app.get('/api/export/account-statement.pdf')
    def export_statement_pdf():
        try:
            result=build()
            if result is None: return jsonify(error='The ledger has changed. Refresh the statement preview.'),409
            return send_file(statement_pdf(result),mimetype='application/pdf',as_attachment=True,
                             download_name=f'MISB_Account_Statement_{result["as_of"]}.pdf')
        except (ValueError,KeyError) as error:
            return jsonify(error=str(error)),400
