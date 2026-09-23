"""New-note screenshot extraction and allocation records."""
import base64
import calendar
import json
import re
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation

from storage import cloud, connect

FIELDS = (
    'loan_code','reference_number','company_id','issuer_name','note_name','product_type','rating',
    'business_description','investment_amount','loan_note_size','allocation_date',
    'disbursal_date','payment_type','term','tenor_type','gross_pa','campaign_start',
    'campaign_end','status','remarks'
)


def allocations():
    con=connect()
    try:
        return [dict(row) for row in con.execute('SELECT * FROM note_allocations ORDER BY updated_at DESC,loan_code')]
    finally:
        con.close()


def allocation_map():
    return {row['loan_code']:row for row in allocations()}


def iso_date(value, label, required=True):
    text=str(value or '').strip()
    if not text and not required: return ''
    try: return date.fromisoformat(text).isoformat()
    except ValueError: raise ValueError(f'{label} must be a valid date') from None


def number(value, label, minimum=0, required=True):
    if value in (None,'') and not required: return None
    try: result=Decimal(str(value))
    except InvalidOperation: raise ValueError(f'{label} must be a number') from None
    if not result.is_finite() or result < Decimal(str(minimum)):
        raise ValueError(f'{label} must be at least {minimum}')
    return float(result)


def normalize_code(value):
    digits=re.search(r'(?:IIF\D*)?(\d{3,6})',str(value or '').upper())
    if not digits: raise ValueError('Reference number must contain an IIF note number')
    return f'IIF-{digits.group(1)}'


def clean_record(data):
    required={
        'issuer_name':'Issuer name','note_name':'Note name','product_type':'Note type',
        'business_description':'Business description','payment_type':'Payment type',
        'tenor_type':'Tenor type'
    }
    raw_reference=str(data.get('reference_number') or data.get('loan_code') or '').strip().upper()
    record={'loan_code':normalize_code(raw_reference),'reference_number':raw_reference}
    for key,label in required.items():
        record[key]=str(data.get(key) or '').strip()
        if not record[key]: raise ValueError(f'{label} is required')
    record['company_id']=int(number(data.get('company_id'),'Company ID',1))
    record['investment_amount']=number(data.get('investment_amount'),'MISB allocation',0.01)
    record['loan_note_size']=number(data.get('loan_note_size'),'Financing amount',0.01)
    record['term']=int(number(data.get('term'),'Note tenure',1))
    record['gross_pa']=number(data.get('gross_pa'),'Profit rate',0)
    # The UI submits percentages (15.60); the simulation stores decimal rates (0.156).
    if record['gross_pa'] > 1: record['gross_pa']/=100
    if record['gross_pa'] > 1: raise ValueError('Profit rate cannot exceed 100%')
    record['allocation_date']=iso_date(data.get('allocation_date'),'Allocation date')
    record['disbursal_date']=iso_date(data.get('disbursal_date'),'Disbursal date')
    record['campaign_start']=iso_date(data.get('campaign_start'),'Campaign start',False)
    record['campaign_end']=iso_date(data.get('campaign_end'),'Campaign end',False)
    if record['campaign_start'] and record['campaign_end'] and record['campaign_end'] < record['campaign_start']:
        raise ValueError('Campaign end cannot be before campaign start')
    if record['tenor_type'] not in ('Days','Months'):
        raise ValueError('Tenor type must be Days or Months')
    if record['payment_type'] not in ('Profit Only','Bullet','Equal Instalment'):
        raise ValueError('Choose a supported payment type')
    record['rating']=str(data.get('rating') or '').strip()
    record['status']=str(data.get('status') or 'Active').strip()
    record['remarks']=str(data.get('remarks') or '').strip()
    return record


def add_months(value, months):
    year=value.year+(value.month-1+months)//12; month=(value.month-1+months)%12+1
    return value.replace(year=year,month=month,day=min(value.day,calendar.monthrange(year,month)[1]))


def final_date(record):
    start=date.fromisoformat(record['disbursal_date']); term=int(record['term'])
    return start+timedelta(days=term) if record['tenor_type']=='Days' else add_months(start,term)


def calculated_values(record):
    investment=float(record['investment_amount']); size=float(record['loan_note_size'])
    rate=float(record['gross_pa']); term=int(record['term'])
    gross=investment*rate*(term/365 if record['tenor_type']=='Days' else term/12)
    net=gross*.8
    return {
        'Investment Exposure':investment/size if size else 0,
        'Final Repayment Date ':final_date(record).isoformat(),
        'Interest Rate p.m (Gross)':rate/12,
        'Interest Rate p.m (Net)':rate/12*.8,
        'Expected Repayment':investment+net,
        'Actual Repayment **':0,'Paid Principal':0,'Unpaid Principal':investment,
        'Gross Profit':gross,'Gross Profit Earned':gross,'Total Gross Profit':gross,
        'Late Payment Charges':0,'SST':0,'Net Profit':net,'Paid Profit':0,'Unpaid Profit':net,
        'Late Profit':0,'Service Fee ':gross-net,
    }


def row_values(headers, record, number_value=''):
    values=calculated_values(record)
    values['Installment']=(values['Gross Profit Earned']/record['term']
                           if record['payment_type']=='Profit Only' and record['term'] else 'Others')
    values.update({
        'No.':number_value,'Loan Code':record['loan_code'],'Company ID':record['company_id'],
        'Issuer Name':record['issuer_name'],'Note Name':record['note_name'],
        'Product Type':record['product_type'],'CTOS/Payment Risk Rating **':record['rating'],
        'Business Description':record['business_description'],'Investment Amount':record['investment_amount'],
        'Loan Note Size':record['loan_note_size'],'Email on Allocation':record['allocation_date'],
        'Email on Disbursement':'Nil','Total exposure against portfolio <20%':'Yes',
        'Payment Type **':record['payment_type'],'Term':record['term'],'Tenor Type':record['tenor_type'],
        'Disbursal Date':record['disbursal_date'],'Loan Status':record['status'],
        'Interest Rate p.a (Gross)':record['gross_pa'],'Early Repayment':'',
        'Early Repayment Date':'','Remarks':record['remarks'],
    })
    return [values.get(header,'') for header in headers]


def save_allocation(record):
    con=connect()
    columns=','.join(FIELDS); marks=','.join('?' for _ in FIELDS)
    updates=','.join(f'{field}=excluded.{field}' for field in FIELDS if field!='loan_code')
    try:
        con.execute(f'''INSERT INTO note_allocations({columns}) VALUES({marks})
            ON CONFLICT(loan_code) DO UPDATE SET {updates},updated_at=CURRENT_TIMESTAMP''',
            tuple(record[field] for field in FIELDS))
        con.commit()
    finally:
        con.close()


def parse_ai_json(answer):
    text=str(answer or '').strip()
    fenced=re.search(r'```(?:json)?\s*(\{.*?\})\s*```',text,re.S|re.I)
    if fenced: text=fenced.group(1)
    else:
        start=text.find('{'); end=text.rfind('}')
        if start>=0 and end>start: text=text[start:end+1]
    result=json.loads(text)
    allowed=('note_name','reference_number','status','note_type','credit_risk_rating',
             'financing_amount','outstanding_balance','outstanding_percentage',
             'profit_rate_pa','note_tenure','tenor_type','campaign_start','campaign_end')
    extracted={key:result.get(key) for key in allowed}
    rate=extracted.get('profit_rate_pa')
    if isinstance(rate,(int,float)) and 0 < rate <= 1:
        extracted['profit_rate_pa']=rate*100
    tenor=str(extracted.get('tenor_type') or '').strip().lower()
    if tenor:
        extracted['tenor_type']='Months' if tenor.startswith('month') else 'Days'
    elif extracted.get('note_tenure'):
        # Cofundr note cards express tenure in days unless another unit is shown.
        extracted['tenor_type']='Days'
    return extracted


def ai_answer(response):
    """Unwrap the result envelope returned by the Python Workers AI binding."""
    if hasattr(response,'to_py'):
        response=response.to_py()
    getter=getattr(response,'get',None)
    if getter:
        nested=getter('result')
        if nested is not None:
            response=nested
    getter=getattr(response,'get',None)
    if getter:
        return getter('answer') or getter('response') or ''
    return getattr(response,'answer','') or getattr(response,'response','')


def month_end(value):
    return value.replace(day=calendar.monthrange(value.year,value.month)[1])


def display_reference(record):
    reference=str(record.get('reference_number') or '').strip()
    if reference and reference != record.get('loan_code'):
        return reference
    campaign=str(record.get('campaign_start') or '')
    if campaign:
        try:
            stamp=date.fromisoformat(campaign).strftime('%d%m%Y')
            digits=re.sub(r'\D','',str(record.get('loan_code') or ''))
            if digits: return f'IIF{digits}-{stamp}'
        except ValueError:
            pass
    return reference or str(record.get('loan_code') or '')


def allocation_email_data(record, request_date, period_end, additional_amount=0,
                          additional_date='', reserve_amount=0, reserve_date=''):
    """Build the allocation email and its cash-position calculation."""
    from app import load_simulation, load_transactions, note_payment_schedule, num, parse_date
    request_day=date.fromisoformat(request_date); end_day=date.fromisoformat(period_end)
    if end_day < request_day:
        raise ValueError('Projection end cannot be before the request date')
    transactions=load_transactions()
    successful=[row for row in transactions if str(row.get('status')).upper()=='SUCCESSFUL'
                and parse_date(row.get('date')) and parse_date(row.get('date'))<=request_day]
    latest=max(successful,key=lambda row:(row.get('date') or '',int(row.get('id') or 0)),default=None)
    available=num(latest.get('current_balance')) if latest else 0
    portfolio,schedule=load_simulation()
    expected=0
    for items in note_payment_schedule(portfolio,schedule,successful).values():
        for item in items:
            due=parse_date(item.get('expected_date'))
            if due and request_day<=due<=end_day and not item.get('is_paid'):
                expected+=max(0,num(item.get('expected_amount')))
    reserve=max(0,float(reserve_amount or 0)); allocation=max(0,float(record['investment_amount']))
    total_expected=available+expected-reserve
    net_available=total_expected-allocation
    required=max(0,allocation-total_expected)
    fmt=lambda value:f'{value:,.2f}'
    slash=lambda value:date.fromisoformat(value).strftime('%d/%m/%Y') if value else ''
    reference=display_reference(record)
    plain=f'''Dear Muamalat Invest Operations Team,

We will be allocating the MiDAS fund into the following investment note:

Note ID:  {reference}
Note Name:  {record['note_name']}
Issuer Name: {record['issuer_name'].upper()}
Amount Allocated: RM {fmt(allocation)}

Total: RM {fmt(allocation)}

Please find attached a screenshot of the note details for your reference. This allocation will be reported under the monthly funding report accordingly and we will notify you once the funds have been disbursed. Attached below is the fund position as at {request_day.day} {request_day.strftime('%B')}.

Thank you.

Regards,
Vinotharan'''
    return {
        'reference_number':reference,'note_name':record['note_name'],'issuer_name':record['issuer_name'],
        'amount':allocation,'request_date':request_date,'period_end':period_end,
        'additional_amount':max(0,float(additional_amount or 0)),'additional_date':additional_date,
        'available':available,'expected':expected,'reserve':reserve,'reserve_date':reserve_date,
        'total_expected':total_expected,'net_available':net_available,'additional_required':required,
        'request_date_display':slash(request_date),'period_end_display':slash(period_end),
        'additional_date_display':slash(additional_date),'reserve_date_display':slash(reserve_date),
        'plain_text':plain,'source_transaction_date':latest.get('date') if latest else None,
    }


def register_allocation_routes(app):
    from flask import jsonify, request

    @app.route('/api/note-allocations',methods=['GET','POST'])
    def note_allocations():
        if request.method=='GET': return jsonify(allocations())
        try:
            record=clean_record(request.get_json() or {});save_allocation(record)
            return jsonify(record)
        except ValueError as error:
            return jsonify(error=str(error)),400

    @app.delete('/api/note-allocations/<path:loan_code>')
    def delete_note_allocation(loan_code):
        code=normalize_code(loan_code);con=connect()
        try: con.execute('DELETE FROM note_allocations WHERE loan_code=?',(code,));con.commit()
        finally: con.close()
        return jsonify(ok=True)

    @app.post('/api/note-allocation/extract')
    def extract_note_allocation():
        image=request.files.get('image')
        if not image: return jsonify(error='Choose a note screenshot first.'),400
        content=image.read()
        if not content or len(content)>3*1024*1024:
            return jsonify(error='The screenshot must be between 1 byte and 3 MB.'),400
        mime=(image.mimetype or '').lower()
        if mime not in ('image/png','image/jpeg','image/webp'):
            return jsonify(error='Use a PNG, JPEG or WebP screenshot.'),400
        env=cloud()
        if not env or not getattr(env,'AI',None):
            return jsonify(error='Screenshot extraction is available on the deployed Cloudflare site.'),503
        question='''Read this financing note card. Return only a JSON object with these keys:
note_name, reference_number, status, note_type, credit_risk_rating, financing_amount,
outstanding_balance, outstanding_percentage, profit_rate_pa, note_tenure, tenor_type,
campaign_start, campaign_end. Use numbers without commas or percent signs. Use YYYY-MM-DD
dates. For a value not visible, use null. Do not infer issuer or business details.'''
        try:
            from pyodide.ffi import run_sync
            payload={'task':'query','image':f'data:{mime};base64,{base64.b64encode(content).decode()}',
                     'question':question,'reasoning':False,'temperature':0,'max_tokens':1000,'stream':False}
            response=run_sync(env.AI.run('@cf/moondream/moondream3.1-9B-A2B',payload))
            extracted=parse_ai_json(ai_answer(response))
            return jsonify(extracted)
        except Exception:
            app.logger.exception('Note screenshot extraction failed')
            return jsonify(error='The screenshot could not be read. Try a clearer image or enter the fields manually.'),422

    @app.get('/api/allocation-email-preview')
    def allocation_email_preview():
        try:
            code=normalize_code(request.args.get('loan_code'))
            record=allocation_map().get(code)
            if not record: return jsonify(error='Save the allocation note before generating its email.'),404
            request_date=iso_date(request.args.get('request_date') or record.get('allocation_date'),'Request date')
            end=iso_date(request.args.get('period_end') or month_end(date.fromisoformat(request_date)).isoformat(),'Projection end')
            additional_date=iso_date(request.args.get('additional_date'),'Additional allocation date',False)
            reserve_date=iso_date(request.args.get('reserve_date'),'Reserve withdrawal date',False)
            additional=number(request.args.get('additional_amount') or 0,'Additional allocation',0)
            reserve=number(request.args.get('reserve_amount') or 0,'Reserve withdrawal',0)
            return jsonify(allocation_email_data(record,request_date,end,additional,additional_date,reserve,reserve_date))
        except ValueError as error:
            return jsonify(error=str(error)),400
