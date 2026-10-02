"""MISB PDF statements: exact decimal ledger amounts and reference artwork."""
import io
import csv
import calendar
from zipfile import BadZipFile
from datetime import datetime, date, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import openpyxl

CENT = Decimal('0.01')
SERVICE_RATE = Decimal('0.20')
SST_RATE = Decimal('0.08')
SST_START = date(2026, 8, 1)
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


def profit_components(net_amount, transaction_date):
    """Split a net profit payout using the SST rules effective on its transaction date."""
    net = money(net_amount)
    taxable = transaction_date >= SST_START
    # RM100 gross - RM20 platform fee - (RM20 * 8%) SST = RM78.40 net.
    divisor = Decimal(1) - SERVICE_RATE * (Decimal(1) + (SST_RATE if taxable else Decimal(0)))
    gross = (net / divisor).quantize(CENT, rounding=ROUND_HALF_UP)
    service = (gross * SERVICE_RATE).quantize(CENT, rounding=ROUND_HALF_UP)
    sst = (service * SST_RATE).quantize(CENT, rounding=ROUND_HALF_UP) if taxable else Decimal(0)
    return gross, service, sst


def statement_data(data, as_of=None, start_date=None, opening_balance=None):
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
            r['amount'] = -abs(r['amount'])
        if action in ('Deposit Approved', 'Withdrawal Approved'):
            r['description'] = action.removesuffix(' Approved')
            if r['note'] == '-':
                r['note'] = '0'
        rows.append(r)
    if not rows:
        raise ValueError('No completed cash transactions exist on or before this cut-off date')
    opening_source = 'First transaction previous balance'
    if start_date:
        try:
            start = date.fromisoformat(start_date)
        except (ValueError, TypeError):
            raise ValueError('Start date must be YYYY-MM-DD') from None
        if start > cutoff:
            raise ValueError('Start date must be on or before the cut-off date')
        prior = [r for r in rows if r['timestamp'].date() < start]
        rows = [r for r in rows if r['timestamp'].date() >= start]
        if opening_balance is not None:
            opening = money(opening_balance)
            opening_source = 'Previous statement closing balance'
        elif prior:
            opening = prior[-1]['current']
            opening_source = 'Last completed transaction before the period'
        elif rows:
            opening = rows[0]['previous']
        else:
            raise ValueError('No opening balance is available for this period')
        if rows and rows[0]['previous'] != opening:
            raise ValueError('The first transaction opening balance does not match the previous closing balance. Check the transaction log for missing entries.')
    else:
        opening = rows[0]['previous']
    breaks = sum(a['current'] != b['previous'] for a,b in zip(rows,rows[1:]))
    if breaks:
        warnings.append(f'{breaks} balance discontinuity/discontinuities found. The PDF preserves the original ledger balances; check that the log is complete.')
    totals = {
        'starting_balance':opening,'ending_balance':rows[-1]['current'] if rows else opening,
        'total_investment':sum((r['amount'] for r in rows if r['description']=='Investment Committed'),Decimal(0)),
        'principal_received':sum((r['amount'] for r in rows if r['description']=='Principal Payout'),Decimal(0)),
        'nett_returns':sum((r['amount'] for r in rows if r['description']=='Profit Payout'),Decimal(0))}
    detailed, gross_total, service_total, sst_total = [], Decimal(0), Decimal(0), Decimal(0)
    for r in rows:
        if r['description'] != 'Profit Payout':
            detailed.append(r)
            continue
        taxable = r['timestamp'].date() >= SST_START
        gross, service, sst = profit_components(r['amount'], r['timestamp'].date())
        common = {'timestamp':r['timestamp'],'order':r['order'],'note':r['note']}
        detailed.append({**common,'description':'Gross Profit','previous':r['previous'],
                         'amount':gross,'current':None})
        detailed.append({**common,'description':'Service Charge','previous':None,
                         'amount':-service,'current':None})
        if taxable:
            detailed.append({**common,'description':'SST','previous':None,
                             'amount':-sst,'current':None})
        detailed.append({**common,'description':'Net Profit','previous':None,
                         'amount':r['amount'],'current':r['current']})
        gross_total += gross; service_total += service; sst_total += sst
    totals.update(total_gross_returns=gross_total, service_fee=service_total, sst=sst_total)
    def value(amount):
        return '' if amount is None else format(amount,'.2f')
    return {'start_date':start_date or '', 'opening_source':opening_source, 'as_of':cutoff.isoformat(),'latest_date':transactions[-1]['timestamp'].date().isoformat(),
        'investor_id':'5490','investor_name':'Amanahraya Trustees Berhad',
        'summary':{k:format(v,'.2f') for k,v in totals.items()},
        'rows':[{'date':r['timestamp'].strftime('%d-%b-%Y %H:%M:%S'),'description':r['description'],
                 'note':r['note'],**{k:value(r[k]) for k in ('previous','amount','current')}} for r in detailed],
        'row_count':len(detailed),'page_count':1+max(0,(len(detailed)-33+57)//58),'warnings':warnings}


def read_statement_csv(data):
    """Read a prepared monthly statement without replacing the full transaction ledger."""
    if len(data)>10*1024*1024: raise ValueError('Statement exceeds 10 MB')
    try:
        table=list(csv.reader(io.StringIO(data.decode('utf-8-sig'))))
    except (UnicodeError,csv.Error):
        raise ValueError('This is not a readable UTF-8 statement CSV') from None
    headers=('Transaction Date','Note ID','Transaction Description',
             'Previous Balance (RM)','Sum Involved (RM)','Current Balance (RM)')
    index=next((i for i,row in enumerate(table) if tuple(row)==headers),None)
    if index is None: raise ValueError('The CSV must use the September monthly statement column headings')
    def amount(value):
        value=value.strip().replace(',','')
        if value.startswith('(') and value.endswith(')'): value='-'+value[1:-1]
        return format(money(value),'.2f') if value else ''
    rows=[];days=[]
    for values in table[index+1:]:
        if not any(value.strip() for value in values): continue
        if len(values)!=6: raise ValueError('Statement CSV rows must have six columns')
        try: day=datetime.strptime(values[0].strip(),'%d-%b-%Y').date()
        except ValueError: raise ValueError('Statement CSV transaction dates must be DD-Mon-YYYY') from None
        days.append(day)
        rows.append(dict(zip(('date','note','description','previous','amount','current'),
            (values[0].strip(),values[1].strip(),values[2].strip(),
             amount(values[3]),amount(values[4]),amount(values[5])))))
    if not rows or len(rows)>10000: raise ValueError('Statement CSV must contain between 1 and 10,000 detail rows')
    labels={'Starting Balance':'starting_balance','Ending Balance':'ending_balance',
            'Total Investment':'total_investment','Total Principal Received':'principal_received',
            'Total Gross Returns Received':'total_gross_returns','Service Fee':'service_fee',
            'SST':'sst','Total Nett Returns Received':'nett_returns'}
    summary={};investor_id='';investor_name=''
    for values in table[:index]:
        if len(values)!=6: continue
        if values[0]=='INVESTOR ID': investor_id=values[2].strip()
        if values[0]=='INVESTOR NAME': investor_name=values[2].strip()
        for label,value in ((values[0],values[1]),(values[3],values[5])):
            if label.strip() in labels: summary[labels[label.strip()]]=amount(value)
    if set(summary)!=set(labels.values()) or any(value=='' for value in summary.values()):
        raise ValueError('Statement CSV is missing account summary amounts')
    if not investor_id or not investor_name: raise ValueError('Statement CSV is missing investor details')
    if investor_id!='5490' or investor_name!='Amanahraya Trustees Berhad':
        raise ValueError('This statement generator supports MISB investor 5490, Amanahraya Trustees Berhad')
    if any(day.year!=days[0].year or day.month!=days[0].month for day in days):
        raise ValueError('Prepared statement CSV must cover one calendar month')
    warnings=[];balance=summary['starting_balance']
    for row in rows:
        if row['previous'] and row['previous']!=balance:
            warnings.append('The prepared statement contains a balance discontinuity. Original balances and row order are preserved.')
            break
        if row['current']: balance=row['current']
    return {'start_date':days[0].replace(day=1).isoformat(),
        'as_of':days[0].replace(day=calendar.monthrange(days[0].year,days[0].month)[1]).isoformat(),
        'investor_id':investor_id,'investor_name':investor_name,'summary':summary,'rows':rows,
        'row_count':len(rows),'page_count':1+max(0,(len(rows)-33+57)//58),
        'warnings':warnings,'mode':'prepared-csv','version':'',
        'opening_source':'Prepared monthly statement; the workspace transaction ledger is unchanged'}


def statement_csv(statement):
    output = io.StringIO(newline='')
    writer = csv.writer(output)
    cutoff=date.fromisoformat(statement['as_of'])
    title=f'ACCOUNT STATEMENT ({MONTHS[cutoff.month-1]} {cutoff.year})' if statement.get('start_date') else f'ACCOUNT STATEMENT (through {statement["as_of"]})'
    writer.writerow((title, '', '', '', '', ''))
    writer.writerow(('INVESTOR ID', '', statement['investor_id'], '', '', ''))
    writer.writerow(('INVESTOR NAME', '', statement['investor_name'], '', '', ''))
    writer.writerow(('ACCOUNT SUMMARY', 'RM', '', 'ACCOUNT SUMMARY', '', 'RM'))
    left=(('Starting Balance','starting_balance'),('Ending Balance','ending_balance'),
          ('Total Investment','total_investment'),('Total Principal Received','principal_received'),
          ('Total Gross Returns Received','total_gross_returns'))
    right=(('Service Fee','service_fee'),('SST','sst'),('Total Nett Returns Received','nett_returns'))
    for index,(label,key) in enumerate(left):
        second=right[index] if index<len(right) else ('','')
        writer.writerow((label,statement['summary'][key],'',second[0],'',statement['summary'].get(second[1],'')))
    writer.writerow(('', '', '', '', '', ''))
    writer.writerow(('DETAILED ACCOUNT STATEMENT', '', '', '', '', ''))
    writer.writerow(('Transaction Date', 'Note ID', 'Transaction Description',
                     'Previous Balance (RM)', 'Sum Involved (RM)', 'Current Balance (RM)'))
    for row in statement['rows']:
        # Keep text literal when the CSV is opened in a spreadsheet.
        cells = [row['date'].split(' ')[0], row['note'], row['description']]
        cells = ["'" + value if value.lstrip().startswith(('=', '+', '-', '@'))
                 and value != '-' else value for value in cells]
        writer.writerow((*cells, row['previous'], row['amount'], row['current']))
    return io.BytesIO(('\ufeff' + output.getvalue()).encode('utf-8'))


def statement_pdf(statement):
    # Lazy imports keep the Worker startup path small.
    from pypdf import PdfReader, PdfWriter
    from pypdf._cmap import get_encoding
    from pypdf.generic import NameObject, DecodedStreamObject
    from statement_template import TEMPLATE
    reader=PdfReader(io.BytesIO(TEMPLATE)); writer=PdfWriter()
    first=writer.add_page(reader.pages[0])
    fonts=first['/Resources']['/Font']

    def font_info(name):
        font=fonts[name].get_object()
        if '/DescendantFonts' not in font:
            return None
        reverse={value:key for key,value in get_encoding(font)[1].items()}
        descendant=font['/DescendantFonts'][0].get_object(); widths={}
        values=descendant.get('/W',[])
        if hasattr(values,'get_object'): values=values.get_object()
        i=0
        while i < len(values):
            start=int(values[i]); following=values[i+1]
            if isinstance(following,list):
                for offset,item in enumerate(following): widths[start+offset]=float(item)
                i+=2
            else:
                end=int(following); item=float(values[i+2])
                for code in range(start,end+1): widths[code]=item
                i+=3
        return reverse,widths,float(descendant.get('/DW',1000))

    font_data={name:font_info(name) for name in ('/F1','/F3')}

    def encoded_text(value,font):
        info=font_data.get(font)
        if info is None:
            raw=value.encode('cp1252')
            f=fonts[font].get_object(); widths=f['/Widths']; start=int(f['/FirstChar'])
            return raw,sum(float(widths[c-start]) if 0<=c-start<len(widths) else 0 for c in raw)
        reverse,widths,default=info; raw=bytearray(); total=0
        for character in value:
            code_text=reverse.get(character)
            if code_text is None:
                raise ValueError(f'The statement font cannot render {character!r}')
            for item in code_text:
                code=ord(item);raw.extend(code.to_bytes(2,'big'));total+=widths.get(code,default)
        return bytes(raw),total

    def width(value,font,size):
        return encoded_text(value,font)[1]*size/1000

    def text(text,x,y,size=6.96,font='/F1',align='left'):
        w=width(text,font,size)
        if align=='right': x-=w
        if align=='center': x-=w/2
        literal=encoded_text(text,font)[0].hex()
        return f'BT 0 g {font} {size} Tf 1 0 0 1 {x:.4f} {792-y:.4f} Tm <{literal}> Tj ET\n'

    cutoff=date.fromisoformat(statement['as_of'])
    title=f'ACCOUNT STATEMENT (YEAR TO DAY- {cutoff.day} {MONTHS[cutoff.month-1]} {cutoff:%y})'
    if statement.get('start_date'):
        title=f'ACCOUNT STATEMENT ({MONTHS[cutoff.month-1]} {cutoff.year})' if statement['start_date'][:7]==statement['as_of'][:7] else f'ACCOUNT STATEMENT ({statement["start_date"]} TO {statement["as_of"]})'
    first_ops=text(title,302.5,190.8,11.4,'/T2','center')
    for key,y in zip(('starting_balance','ending_balance','total_investment','principal_received','total_gross_returns'),(261.96,273.60,285.24,296.88,308.52)):
        first_ops+=text(format(Decimal(statement['summary'][key]),',.2f'),312.95,y,7.56,font='/F1',align='right')
    for key,y in zip(('service_fee','sst','nett_returns'),(261.96,273.60,285.24)):
        first_ops+=text(format(Decimal(statement['summary'][key]),',.2f'),550.33,y,7.56,font='/F1',align='right')
    # Draw the new Date / Note ID / Description heading over the previous artwork.
    first_ops+='0.02 0.12 0.36 rg 50.964 438.36 309.766 19.0 re f\n'
    for label,x in (('Transaction Date',100),('Note ID',178),('Transaction Description',270)):
        first_ops+=text(label,x,348.6,6.96,'/F1','center').replace('BT 0 g','BT 1 g')
    all_rows=statement['rows']; groups=[all_rows[:33]]+[all_rows[i:i+58] for i in range(33,len(all_rows),58)]
    columns=(50.964,151.0,205.0,360.73,418.45,487.21,553.5)

    def display_amount(value):
        if value == '': return ''
        amount=Decimal(value)
        return f'({abs(amount):,.2f})' if amount < 0 else f'{amount:,.2f}'
    for page_index,rows in enumerate(groups):
        page=first if page_index==0 else writer.add_blank_page(width=612,height=792)
        if page_index: page[NameObject('/Resources')]=first['/Resources']
        ops=first_ops if page_index==0 else ''
        top=353.64 if page_index==0 else 54.24
        bottom=top+len(rows)*11.64
        ops+='0 G 0.6 w\n'
        for x in columns:
            ops+=f'{x} {792-top:.4f} m {x} {792-bottom:.4f} l S\n'
        for i in range(len(rows)+1):
            y=792-top-i*11.64
            ops+=f'{columns[0]} {y:.4f} m {columns[-1]} {y:.4f} l S\n'
        for i,row in enumerate(rows):
            y=(360.60 if page_index==0 else 62.52)+11.64*i
            for value,left,right in zip((row['date'].split(' ')[0],row['note'],row['description']),columns[:3],columns[1:4]):
                size=min(6.96,6.96*(right-left-3)/max(1,width(value,'/F1',6.96)))
                ops+=text(value,(left+right)/2,y,size,font='/F1',align='center')
            for key,right in zip(('previous','amount','current'),(409.35,486.85,550.40)):
                value=display_amount(row[key]); font='/F3' if value.startswith('(') else '/F1'
                ops+=text(value,right,y,font=font,align='right')
        old=page.get_contents()
        stream=DecodedStreamObject()
        stream.set_data((b'q\n'+old.get_data()+b'\nQ\n' if old is not None else b'')+ops.encode('ascii'))
        page[NameObject('/Contents')]=stream
        page.compress_content_streams()
    writer.add_metadata({'/Title':f'MISB Account Statement through {statement["as_of"]}', '/Author':'Crowd Sense Sdn Bhd'})
    output=io.BytesIO();writer.write(output);output.seek(0)
    return output


def register_statement_routes(app):
    from flask import request, jsonify, send_file, g
    from storage import (read_source, source_rows, save_source, source_exists,
                         save_parsed_source, audit_event)
    from uploads import inspect_workbook

    def prepare(data, params):
        cutoff = params.get('as_of')
        if params.get('mode') != 'monthly':
            return statement_data(data, cutoff)
        from storage import statement_report_runs
        full = statement_data(data, cutoff)
        end = full['as_of']
        start = params.get('start_date')
        previous = sorted((r for r in statement_report_runs(500) if r['as_of'] < (start or end)),
                          key=lambda r: (r['as_of'], r['id']), reverse=True)
        opening = None
        if not start:
            if previous:
                start = (date.fromisoformat(previous[0]['as_of']) + timedelta(days=1)).isoformat()
                opening = previous[0]['closing_balance']
            else:
                anchor = '2026-09-02'
                start = '2026-09-03' if end > anchor else end[:8] + '01'
        elif previous and (date.fromisoformat(previous[0]['as_of']) + timedelta(days=1)).isoformat() == start:
            opening = previous[0]['closing_balance']
        result = statement_data(data, end, start, opening)
        result['mode'] = 'monthly'
        return result

    def build():
        if not source_exists('transactions'):
            raise ValueError('Upload a transaction log first')
        expected=request.args.get('version')
        if expected and expected!=source_rows()['transactions']['object_key']:
            return None
        result=prepare(read_source('transactions').read(),request.args)
        result['version']=source_rows()['transactions']['object_key']
        return result

    @app.route('/api/account-statement',methods=['GET','POST'])
    def preview_statement():
        try:
            file=request.files.get('file')
            if file and file.filename.lower().endswith('.csv'):
                return jsonify(read_statement_csv(file.read()))
            if request.method=='POST':
                file=request.files.get('file')
                if not file or not file.filename.lower().endswith('.xlsx'):
                    raise ValueError('Choose a transaction log in .xlsx format')
                data=file.read()
                from app import load_transactions
                loaded=load_transactions() if source_exists('transactions') else []
                current=list(getattr(g,'transactions_base',loaded))
                quality=inspect_workbook('transactions',data,current)
                overwrite=request.form.get('overwrite') in ('1','true','yes') and quality.get('can_overwrite',False)
                if not quality.get('can_import') and not overwrite:
                    return jsonify(error='The transaction workbook failed data-quality checks.',
                                   issues=quality.get('issues',[])),422
                if ((quality.get('comparison') or {}).get('requires_acknowledgement')
                        and request.form.get('allow_regression') not in ('1','true','yes')):
                    return jsonify(error='This ledger is older or removes live transactions. Review and import it from Data Sources if that rollback is intentional.',
                                   comparison=quality['comparison']),409
                result=prepare(data,request.form)
                if overwrite:
                    result['warnings'].extend(f"Imported ledger warning: {item['message']} ({item['count']})"
                                              for item in quality.get('issues',[]) if item['level']=='error')
                save_source('transactions',data,result['latest_date'],file.filename)
                g.pop('transactions_base',None);g.pop('transactions_effective',None)
                load_transactions()
                save_parsed_source('transactions',source_rows()['transactions']['object_key'],g.transactions_base)
                audit_event('Imported statement ledger','source','transactions',{'as_of':result['latest_date'],
                    'quality_overridden':bool(overwrite),'quality_issues':quality.get('issues',[]) if overwrite else [],
                    'regression_acknowledged':bool((quality.get('comparison') or {}).get('requires_acknowledgement'))})
                result['version']=source_rows()['transactions']['object_key']
            else:
                result=build()
            if result is None: return jsonify(error='The ledger has changed. Refresh the statement preview.'),409
            return jsonify(result)
        except (BadZipFile,openpyxl.utils.exceptions.InvalidFileException):
            return jsonify(error='This file is not a readable Excel workbook.'),400
        except (ValueError,KeyError,StopIteration) as error:
            return jsonify(error=str(error)),400

    @app.route('/api/export/account-statement.csv',methods=['GET','POST'])
    def export_statement_csv():
        try:
            result = read_statement_csv(request.files['file'].read()) if request.method=='POST' and 'file' in request.files else build()
            if result is None:
                return jsonify(error='The ledger has changed. Refresh the statement preview.'),409
            period = (result['start_date'] + '_to_') if result.get('start_date') else ''
            filename = f'MISB_Account_Statement_{period}{result["as_of"]}.csv'
            audit_event('Downloaded account statement CSV', 'report', result['as_of'], {
                'filename':filename, 'source_version':result['version'], 'rows':result['row_count']})
            return send_file(statement_csv(result), mimetype='text/csv',
                             as_attachment=True, download_name=filename)
        except (ValueError,KeyError) as error:
            return jsonify(error=str(error)),400

    @app.route('/api/export/account-statement.pdf',methods=['GET','POST'])
    def export_statement_pdf():
        try:
            if request.method=='POST' and 'file' in request.files:
                result=read_statement_csv(request.files['file'].read())
                filename=f'MISB_Account_Statement_{result["start_date"]}_to_{result["as_of"]}.pdf'
                return send_file(statement_pdf(result),mimetype='application/pdf',as_attachment=True,download_name=filename)
            result=build()
            if result is None: return jsonify(error='The ledger has changed. Refresh the statement preview.'),409
            filename=f'MISB_Account_Statement_{result.get("start_date") + "_to_" if result.get("start_date") else ""}{result["as_of"]}.pdf';pdf=statement_pdf(result)
            from storage import save_statement_report_run
            summary=result.get('summary',{})
            record=save_statement_report_run(result['as_of'],result['version'],filename,{
                'start_date':result.get('start_date',''),'row_count':result.get('row_count',0),'page_count':result.get('page_count',0),
                'opening_balance':summary.get('starting_balance',0),'closing_balance':summary.get('ending_balance',0),
                'gross_returns':summary.get('total_gross_returns',0)})
            audit_event('Generated account statement' if record['created'] else 'Downloaded existing account statement',
                        'report',result['as_of'],{
                'history_id':record['id'],'filename':filename,'source_version':result['version'],
                'rows':result.get('row_count',0)})
            response=send_file(pdf,mimetype='application/pdf',as_attachment=True,download_name=filename)
            response.headers['X-Statement-History-ID']=str(record['id'])
            response.headers['X-Statement-History-New']='1' if record['created'] else '0'
            return response
        except (ValueError,KeyError) as error:
            return jsonify(error=str(error)),400

    @app.get('/api/account-statement-runs')
    def account_statement_history():
        from storage import statement_report_runs
        return jsonify(statement_report_runs(request.args.get('limit',20)))

    @app.get('/api/account-statement-runs.csv')
    def account_statement_register():
        from storage import statement_report_runs
        output=io.StringIO(newline='')
        writer=csv.writer(output)
        writer.writerow(('Generated at','Statement through','PDF filename','Rows','Pages','Opening balance (RM)',
                         'Closing balance (RM)','Gross returns (RM)','Source workbook','Source SHA-256',
                         'Source size (bytes)','Source available','Last verified at','Verification result',
                         'Mismatched fields','Period starts'))
        for run in statement_report_runs(500):
            verification=('Never verified' if run['last_verification_matches'] is None else
                          ('Matches' if run['last_verification_matches'] else 'Review mismatch'))
            writer.writerow((run['created_at'],run['as_of'],run['filename'],run['row_count'],run['page_count'],
                             f"{run['opening_balance']:.2f}",f"{run['closing_balance']:.2f}",
                             f"{run['gross_returns']:.2f}",run['source_filename'],run['source_sha256'],
                             run['source_byte_size'],'Yes' if run['source_available'] else 'No',
                             run['last_verified_at'] or '',verification,
                             ', '.join(run['last_verification_mismatches']),run.get('start_date','')))
        data=io.BytesIO(('\ufeff'+output.getvalue()).encode('utf-8'))
        return send_file(data,mimetype='text/csv',as_attachment=True,
                         download_name='MISB_Account_Statement_Register.csv')

    @app.post('/api/account-statement-runs/<int:run_id>/verify')
    def verify_account_statement(run_id):
        from storage import statement_report_run,read_source_version,save_statement_verification
        run=statement_report_run(run_id)
        if not run:
            return jsonify(error='That statement history record was not found.'),404
        try:
            source,_=read_source_version('transactions',run['source_object_key'])
            result=statement_data(source.read(),run['as_of'],run.get('start_date') or None,
                                  run['opening_balance'] if run.get('start_date') else None)
            summary=result.get('summary',{})
            values={
                'Rows':(int(run['row_count']),int(result['row_count'])),
                'Pages':(int(run['page_count']),int(result['page_count'])),
                'Opening balance':(f"{Decimal(str(run['opening_balance'])):.2f}",summary['starting_balance']),
                'Closing balance':(f"{Decimal(str(run['closing_balance'])):.2f}",summary['ending_balance']),
                'Gross returns':(f"{Decimal(str(run['gross_returns'])):.2f}",summary['total_gross_returns'])}
            checks={name:{'recorded':recorded,'calculated':calculated,
                          'match':recorded==calculated}
                    for name,(recorded,calculated) in values.items()}
            matches=all(check['match'] for check in checks.values())
            mismatches=[name for name,check in checks.items() if not check['match']]
            verification=save_statement_verification(run_id,matches,mismatches)
            audit_event('Verified account statement','report',run['as_of'],{
                'history_id':run_id,'matches':matches,
                'mismatches':mismatches,
                'source_version':run['source_object_key']})
            return jsonify({'history_id':run_id,'matches':matches,'checks':checks,
                            'source_sha256':run['source_sha256'],
                            'verified_at':verification['verified_at']})
        except (ValueError,KeyError) as error:
            return jsonify(error=str(error)),409

    @app.get('/api/account-statement-runs/<int:run_id>/pdf')
    def reproduce_account_statement(run_id):
        from storage import statement_report_run,read_source_version
        run=statement_report_run(run_id)
        if not run: return jsonify(error='That statement history record was not found.'),404
        try:
            source,_=read_source_version('transactions',run['source_object_key'])
            result=statement_data(source.read(),run['as_of'],run.get('start_date') or None,
                                  run['opening_balance'] if run.get('start_date') else None)
            audit_event('Recreated account statement','report',run['as_of'],{
                'history_id':run_id,'filename':run['filename'],'source_version':run['source_object_key']})
            response=send_file(statement_pdf(result),mimetype='application/pdf',as_attachment=True,
                               download_name=run['filename'])
            if run.get('source_sha256'):
                response.headers['X-Statement-Source-SHA256']=run['source_sha256']
            return response
        except (ValueError,KeyError) as error:
            return jsonify(error=str(error)),400
