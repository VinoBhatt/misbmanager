from flask import Flask, jsonify, request, send_file, g
from pathlib import Path
from datetime import datetime, date, timedelta
import openpyxl, json, re, io, csv, calendar
from storage import connect, read_source, source_exists, source_rows, save_source

app = Flask(__name__, static_folder=None)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024


def num(v):
    try: return float(v or 0)
    except: return 0.0

def iso(v):
    if isinstance(v, (datetime, date)): return v.strftime('%Y-%m-%d')
    if v is None: return None
    s=str(v).strip()
    for fmt in ('%d-%b-%Y %H:%M:%S','%d/%m/%Y','%Y-%m-%d %H:%M:%S','%Y-%m-%d'):
        try: return datetime.strptime(s,fmt).strftime('%Y-%m-%d')
        except: pass
    return s

def normalize_note_code(v):
    if v is None: return ''
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try: return f'IIF-{int(v)}'
        except: pass
    s=str(v).replace('**','').strip().upper()
    m=re.search(r'(?:IIF[-\s]*)?(\d{3,6})', s)
    return f'IIF-{m.group(1)}' if m else s

def parse_date(v):
    s=iso(v)
    try: return datetime.strptime(s,'%Y-%m-%d').date() if s else None
    except: return None

def add_months(d, months):
    if not d: return None
    y=d.year+(d.month-1+months)//12; m=(d.month-1+months)%12+1
    return date(y,m,min(d.day,calendar.monthrange(y,m)[1]))

def month_key(d): return d.strftime('%Y-%m') if d else ''

def month_label(k):
    try: return datetime.strptime(k,'%Y-%m').strftime('%b %Y')
    except: return k

def get_settings():
    con=connect()
    d={r['key']:r['value'] for r in con.execute('SELECT key,value FROM settings')}; con.close()
    return {'issuer_limit':num(d.get('issuer_limit',2000000))}

def load_transactions():
    if not source_exists('transactions'): return []
    wb=openpyxl.load_workbook(read_source('transactions'),data_only=True,read_only=True)
    ws=wb[wb.sheetnames[0]]
    headers=[c.value for c in next(ws.iter_rows(min_row=2,max_row=2))]
    out=[]
    for vals in ws.iter_rows(min_row=3,values_only=True):
        if not any(v is not None for v in vals): continue
        r=dict(zip(headers,vals))
        out.append({
            'id': r.get('ID'), 'action': r.get('Action') or '', 'amount': num(r.get('Sum Involved(MYR)')),
            'previous_balance': num(r.get('Previous Balance(MYR)')), 'current_balance': num(r.get('Current Balance(MYR)')),
            'note': normalize_note_code(r.get('Note #')), 'transaction_no': r.get('Transaction No') or '',
            'status': r.get('Status') or '', 'date_raw': str(r.get('Date') or ''), 'date': iso(r.get('Date')), 'pay_via': r.get('Pay Via') or ''
        })
    return out

def load_simulation():
    if not source_exists('simulation'): return [], []
    wb=openpyxl.load_workbook(read_source('simulation'),data_only=True,read_only=True)
    ws=wb['Query result'] if 'Query result' in wb.sheetnames else wb[wb.sheetnames[0]]
    headers=[c.value for c in next(ws.iter_rows(min_row=1,max_row=1))]
    portfolio=[]
    for vals in ws.iter_rows(min_row=2,values_only=True):
        if not vals or vals[0] is None or not vals[1]: continue
        r=dict(zip(headers,vals))
        portfolio.append({
            'loan_code': normalize_note_code(r.get('Loan Code')),
            'issuer': r.get('Issuer Name') or '', 'note_name': r.get('Note Name') or '', 'product': r.get('Product Type') or '',
            'rating': r.get('CTOS/Payment Risk Rating **') or '', 'business': r.get('Business Description') or '',
            'investment_amount': num(r.get('Investment Amount')), 'note_size': num(r.get('Loan Note Size')),
            'exposure': num(r.get('Investment Exposure')), 'payment_type': r.get('Payment Type **') or '', 'term': r.get('Term') or '',
            'disbursal_date': iso(r.get('Disbursal Date')), 'final_date': iso(r.get('Final Repayment Date ')),
            'loan_status': r.get('Loan Status') or '', 'gross_pa': num(r.get('Interest Rate p.a (Gross)')),
            'gross_pm': num(r.get('Interest Rate p.m (Gross)')), 'net_pm': num(r.get('Interest Rate p.m (Net)')),
            'expected_repayment': num(r.get('Expected Repayment')), 'actual_repayment': num(r.get('Actual Repayment **')),
            'paid_principal': num(r.get('Paid Principal')), 'unpaid_principal': num(r.get('Unpaid Principal')),
            'gross_profit': num(r.get('Gross Profit')), 'net_profit': num(r.get('Net Profit')), 'paid_profit': num(r.get('Paid Profit')),
            'unpaid_profit': num(r.get('Unpaid Profit')), 'late_profit': num(r.get('Late Profit')), 'service_fee': num(r.get('Service Fee ')),
            'early_repayment': r.get('Early Repayment') or '', 'early_date': iso(r.get('Early Repayment Date')), 'remarks': r.get('Remarks') or ''
        })
    schedule=[]
    if 'Sheet1' in wb.sheetnames:
        s=wb['Sheet1']; months=[]
        for c in range(7,14):
            m=s.cell(4,c).value
            if m: months.append((c,str(m)))
        for row in range(5,s.max_row+1):
            code=s.cell(row,1).value
            if not code: continue
            item={'loan_code':normalize_note_code(code),'monthly':num(s.cell(row,2).value),'gross_nett':s.cell(row,3).value or '',
                  'tenure':s.cell(row,4).value or '', 'final_date':iso(s.cell(row,5).value),'payment_type':s.cell(row,6).value or '', 'months':{}}
            for c,m in months:
                v=s.cell(row,c).value
                if isinstance(v,(int,float)): item['months'][m]=float(v)
                elif v: item['months'][m]=str(v)
            schedule.append(item)
    return portfolio, schedule


def load_projection_history():
    """Read realised historical receipt entries from the legacy MIDAS daily projection model.

    The legacy workbook models each day as: closing balance = opening balance + cash in - cash out.
    We use only the historical/realised portion to backfill old note repayment months; the live transaction
    ledger always wins when the same note/payment is present there.
    """
    if not source_exists('projections'): return {'cutoff':None,'events':[]}
    wb=openpyxl.load_workbook(read_source('projections'),data_only=True,read_only=True)
    # Determine the last explicitly historical month from Sheet1's "as at" blocks.
    cutoff=None
    if 'Sheet1' in wb.sheetnames:
        ws1=wb['Sheet1']
        for row in ws1.iter_rows(values_only=True):
            for v in row:
                if isinstance(v,str):
                    m=re.search(r'as\s+at\s+(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?',v,re.I)
                    if m:
                        yy=int(m.group(3)) if m.group(3) else 2026
                        if yy<100: yy+=2000
                        try:
                            d=date(yy,int(m.group(2)),int(m.group(1)))
                            cutoff=max(cutoff,d) if cutoff else d
                        except: pass
    if cutoff is None: cutoff=date(2026,6,30)
    events=[]
    if 'Sheet2' in wb.sheetnames:
        ws=wb['Sheet2']
        # Date headings are on row 3; Cash In detail occupies rows 8:51 in this workbook.
        # Some Excel/OpenPyXL combinations can expose max_column/max_row as None
        # for a partially read or effectively empty worksheet.  Treat that source
        # as having no historical schedule instead of taking down the whole API.
        max_col = getattr(ws, 'max_column', 0) or 0
        max_row = getattr(ws, 'max_row', 0) or 0
        if max_col < 4 or max_row < 8:
            return {'cutoff': cutoff.isoformat() if cutoff else None, 'events': []}
        dates={c:parse_date(ws.cell(3,c).value) for c in range(4,max_col+1)}
        current_label=''
        for r in range(8,min(52,max_row+1)):
            label=ws.cell(r,2).value
            typ=str(ws.cell(r,3).value or '').strip().title()
            if typ not in ('Profit','Principal'): continue
            # Principal rows carry the note label; the following Profit row often carries only timing text.
            if typ=='Principal' and label: current_label=str(label)
            elif label and 'IIF' in str(label).upper(): current_label=str(label)
            code=normalize_note_code(current_label)
            if not code.startswith('IIF-'): continue
            for c,d in dates.items():
                if not d or d>cutoff: continue
                amt=ws.cell(r,c).value
                if isinstance(amt,(int,float)) and abs(float(amt))>0.005:
                    events.append({'loan_code':code,'type':typ,'date':d,'amount':float(amt),'source':'MIDAS Projections'})
    return {'cutoff':cutoff.isoformat() if cutoff else None,'events':events}

def ledger_payouts(transactions, code, action, exclude_special_profit=False):
    """Return successful payout totals by transaction date for a note."""
    vals={}
    for t in transactions:
        if t.get('note')!=code or t.get('action')!=action or str(t.get('status','')).upper()!='SUCCESSFUL':
            continue
        detail=str(t.get('transaction_no') or '').lower()
        if exclude_special_profit and ('deferred profit' in detail or 'tawidh' in detail or "ta'widh" in detail):
            continue
        d=parse_date(t.get('date'))
        if d: vals[d]=vals.get(d,0)+num(t.get('amount'))
    return sorted(vals.items())


def full_principal_settlement_date(r, transactions):
    """Detect the date on which cumulative principal payouts fully settled the MISB allocation."""
    target=max(0,num(r.get('investment_amount')))
    payouts=ledger_payouts(transactions,r['loan_code'],'Principal Payout')
    if not payouts: return None
    running=0
    tolerance=max(1.0,target*0.00001)
    for d,amt in payouts:
        running+=amt
        if target and running >= target-tolerance:
            return d
    # The simulation's completed/early-settlement flag is useful when rounding or legacy rows prevent exact equality.
    if str(r.get('loan_status') or '').lower()=='completed' and r.get('early_date'):
        return parse_date(r.get('early_date')) or payouts[-1][0]
    return None


def apply_early_settlement_to_profit_events(r, events, settlement):
    """Close future profit periods after an early full-principal settlement and insert the pro-rated final profit period.

    Pro-rata rule: outstanding principal × monthly rate × (days from last instalment date to settlement / 30).
    For Profit Only notes the outstanding principal is the original MISB allocation. For Bullet notes the period starts
    at disbursement. Equal-instalment notes use the scheduled outstanding principal after periods already due.
    """
    if not settlement: return events
    disb=parse_date(r.get('disbursal_date')); final=parse_date(r.get('final_date'))
    if not disb or not final or settlement >= final: return events
    ptype=str(r.get('payment_type') or '').lower(); rate=max(0,num(r.get('net_pm')))
    original=max(0,num(r.get('investment_amount')))
    if 'bullet' in ptype:
        days=max(0,(settlement-disb).days)
        expected=original*rate*days/30
        return [{'due_date':settlement,'expected':expected,'early_settlement':True,'original_due_date':final,'prorata_days':days}]

    kept=[dict(e) for e in events if e['due_date'] <= settlement]
    if any(e['due_date']==settlement for e in kept):
        kept[-1]['early_settlement']=True; kept[-1]['original_due_date']=kept[-1]['due_date']; kept[-1]['prorata_days']=30
        return kept
    previous_due=max([e['due_date'] for e in events if e['due_date'] < settlement], default=disb)
    next_due=min([e['due_date'] for e in events if e['due_date'] > settlement], default=final)
    days=max(0,(settlement-previous_due).days)
    outstanding=original
    if 'equal instalment' in ptype:
        try: term=max(1,int(float(r.get('term') or 1)))
        except: term=1
        periods_due=sum(1 for e in events if e['due_date'] <= previous_due)
        outstanding=max(0,original-(original/term)*periods_due)
    expected=outstanding*rate*days/30
    kept.append({'due_date':settlement,'expected':expected,'early_settlement':True,'original_due_date':next_due,'prorata_days':days})
    kept.sort(key=lambda e:e['due_date'])
    return kept


def match_profit_payouts(events, payouts):
    """Allocate ledger profit cash across scheduled periods, including multiple periods paid on one transaction date."""
    eps=0.005
    idx=0
    for pd,pool0 in payouts:
        pool=max(0,num(pool0))
        while pool>eps and idx<len(events):
            e=events[idx]
            expected=max(0,num(e.get('expected')))
            # Normal periods consume their scheduled amount. The final/early-settlement period consumes the remainder.
            if e.get('early_settlement') or idx==len(events)-1:
                actual=pool
            else:
                actual=min(pool, expected if expected>eps else pool)
            e.update({'paid':True,'paid_date':pd,'actual':actual})
            pool-=actual; idx+=1
    for i,e in enumerate(events):
        if 'paid' not in e: e.update({'paid':False,'paid_date':None,'actual':0})
    return events



def reconcile_profit_events(r, standard_events, settlement, payouts):
    """Reconcile scheduled profit with ledger cash. Early settlement reserves the settlement-date profit as
    the final pro-rated period, and determines the pro-rata start from the last instalment actually covered.
    """
    if not settlement:
        return match_profit_payouts([dict(e) for e in standard_events],payouts)
    disb=parse_date(r.get('disbursal_date')); final=parse_date(r.get('final_date'))
    if not disb or not final or settlement >= final:
        return match_profit_payouts([dict(e) for e in standard_events],payouts)

    before=[(d,a) for d,a in payouts if d < settlement]
    on_settlement=sum(a for d,a in payouts if d == settlement)
    after=[(d,a) for d,a in payouts if d > settlement]
    regular=[dict(e) for e in standard_events if e['due_date'] < settlement]
    regular=match_profit_payouts(regular,before)
    # Keep only regular periods actually satisfied before settlement. Any unsatisfied later schedule is extinguished.
    regular_paid=[e for e in regular if e.get('paid')]
    last_instalment=max([e['due_date'] for e in regular_paid],default=disb)
    ptype=str(r.get('payment_type') or '').lower()
    if 'bullet' in ptype: last_instalment=disb
    days=max(0,(settlement-last_instalment).days)
    outstanding=max(0,num(r.get('investment_amount')))
    if 'equal instalment' in ptype:
        try: term=max(1,int(float(r.get('term') or 1)))
        except: term=1
        outstanding=max(0,outstanding-(outstanding/term)*len(regular_paid))
    net_rate=max(0,num(r.get('net_pm'))); gross_rate=max(0,num(r.get('gross_pm')))
    early={'due_date':settlement,'expected':outstanding*net_rate*days/30,'expected_gross':outstanding*gross_rate*days/30,
           'early_settlement':True,'original_due_date':final,'prorata_days':days,'prorata_from':last_instalment}
    if on_settlement>0:
        early.update({'paid':True,'paid_date':settlement,'actual':on_settlement})
    elif after:
        # Allow a short processing delay after the principal settlement; the principal date remains the economic settlement date.
        pd,amt=after[0]
        if (pd-settlement).days <= 3: early.update({'paid':True,'paid_date':pd,'actual':amt})
        else: early.update({'paid':False,'paid_date':None,'actual':0})
    else:
        early.update({'paid':False,'paid_date':None,'actual':0})
    return regular_paid+[early]

def derived_profit_schedule(portfolio, transactions, projection_history=None):
    """Build profit periods from note terms, then reconcile them to the transaction ledger.
    A full early principal settlement closes all later periods and creates a final pro-rated profit period.
    """
    rows=[]; all_months=set()
    for r in portfolio:
        code=r['loan_code']; disb=parse_date(r.get('disbursal_date')); final=parse_date(r.get('final_date'))
        try: term=max(0,int(float(r.get('term') or 0)))
        except: term=0
        ptype=str(r.get('payment_type') or '').strip(); pl=ptype.lower(); events=[]
        if disb and term:
            if 'bullet' in pl:
                due=final or add_months(disb,term)
                expected=max(0,num(r.get('net_profit'))) or max(0,num(r.get('investment_amount'))*num(r.get('net_pm'))*term)
                if due: events.append({'due_date':due,'expected':expected})
            else:
                monthly=max(0,num(r.get('investment_amount'))*num(r.get('net_pm')))
                for i in range(1,term+1):
                    due=add_months(disb,i)
                    if final and i==term: due=final
                    if due: events.append({'due_date':due,'expected':monthly})
        elif final:
            events=[{'due_date':final,'expected':max(0,num(r.get('unpaid_profit')) or num(r.get('net_profit')))}]

        settlement=full_principal_settlement_date(r,transactions)
        payouts=ledger_payouts(transactions,code,'Profit Payout',True)
        # Backfill realised legacy periods only when the transaction ledger has no matching receipt.
        hist=[]
        if projection_history:
            hist=[(x['date'],x['amount']) for x in projection_history.get('events',[]) if x['loan_code']==code and x['type']=='Profit']
        if hist:
            ledger_dates={d for d,_ in payouts}
            for hd,ha in hist:
                if not any(abs((hd-ld).days)<=3 for ld in ledger_dates): payouts.append((hd,ha))
            payouts.sort(key=lambda z:z[0])
        events=reconcile_profit_events(r,events,settlement,payouts)
        for event in events:
            due=event['due_date']; all_months.add(month_key(due))
            if event['paid']:
                delta=(event['paid_date']-due).days
                event['status']='Paid'; event['timing']='early' if delta<0 else ('late' if delta>0 else 'on_time'); event['days_variance']=delta
            else:
                event['status']='Unpaid' if due < date.today() else 'Expected'; event['timing']='overdue' if due < date.today() else 'future'; event['days_variance']=None
        cells={}
        for e in events:
            k=month_key(e['due_date']); cells[k]={
                'due_date':e['due_date'].isoformat(),'expected':e['expected'],'paid':e['paid'],
                'paid_date':e['paid_date'].isoformat() if e['paid_date'] else '', 'actual':e['actual'],
                'status':e['status'],'timing':e['timing'],'days_variance':e['days_variance'],
                'early_settlement':bool(e.get('early_settlement')),
                'original_due_date':e.get('original_due_date').isoformat() if e.get('original_due_date') else '',
                'prorata_days':e.get('prorata_days'), 'prorata_from':e.get('prorata_from').isoformat() if e.get('prorata_from') else '',
                'expected_gross':e.get('expected_gross'),
                'source': ('MIDAS Projections' if projection_history and e.get('paid_date') and any(x['loan_code']==code and x['type']=='Profit' and x['date']==e.get('paid_date') for x in projection_history.get('events',[])) else ('Transaction ledger' if e.get('paid') else ''))
            }
        rows.append({'loan_code':code,'payment_type':ptype,'disbursal_date':r.get('disbursal_date'),'term':term,
                     'final_date':r.get('final_date'),'loan_status':r.get('loan_status'),'settlement_date':settlement.isoformat() if settlement else '', 'months':cells})
    months=sorted(all_months)
    return {'months':[{'key':k,'label':month_label(k)} for k in months], 'rows':rows}

def monitoring_map():
    con=connect()
    rows={r['loan_code']:dict(r) for r in con.execute('SELECT * FROM monitoring')}
    con.close(); return rows

def is_active(r): return bool(r['loan_status']) and r['loan_status'].lower() != 'completed'

def payment_marks_map():
    con=connect()
    rows=[dict(r) for r in con.execute('SELECT * FROM payment_marks ORDER BY expected_date,event_key')]
    con.close(); return {(r['loan_code'],r['event_key']):r for r in rows}

def cashflow_plan_rows():
    con=connect()
    rows=[dict(r) for r in con.execute('SELECT * FROM cashflow_plan ORDER BY flow_date,id')]
    con.close(); return rows

def note_payment_schedule(pf, schedule, transactions=None):
    marks=payment_marks_map(); transactions=transactions or []; out={}
    for r in pf:
        code=r['loan_code']; items=[]; disb=parse_date(r.get('disbursal_date')); final=parse_date(r.get('final_date'))
        try: term=max(0,int(float(r.get('term') or 0)))
        except: term=0
        ptype=str(r.get('payment_type') or '').lower(); profit_due=[]; principal_due=[]
        if disb and term:
            if 'bullet' in ptype:
                due=final or add_months(disb,term)
                if due:
                    profit_due=[{'due_date':due,'expected':max(0,num(r.get('net_profit'))) or max(0,num(r.get('investment_amount'))*num(r.get('net_pm'))*term)}]
                    principal_due=[{'due_date':due,'expected':max(0,num(r.get('investment_amount')))}]
            else:
                mp=max(0,num(r.get('investment_amount'))*num(r.get('net_pm')))
                for i in range(1,term+1):
                    due=final if (final and i==term) else add_months(disb,i)
                    if due: profit_due.append({'due_date':due,'expected':mp})
                if 'equal instalment' in ptype:
                    princ=max(0,num(r.get('investment_amount'))/term) if term else 0
                    principal_due=[{'due_date':x['due_date'],'expected':princ} for x in profit_due]
                elif final:
                    principal_due=[{'due_date':final,'expected':max(0,num(r.get('investment_amount')))}]
        elif final:
            profit_due=[{'due_date':final,'expected':max(0,num(r.get('unpaid_profit')) or num(r.get('net_profit')))}]
            principal_due=[{'due_date':final,'expected':max(0,num(r.get('unpaid_principal')) or num(r.get('investment_amount')))}]

        settlement=full_principal_settlement_date(r,transactions)
        profit_due=reconcile_profit_events(r,profit_due,settlement,ledger_payouts(transactions,code,'Profit Payout',True))

        # Once principal is fully settled, no later principal instalments remain receivable.
        if settlement and final and settlement < final:
            if 'equal instalment' in ptype:
                principal_due=[x for x in principal_due if x['due_date'] < settlement]
                paid_before=sum(a for d,a in ledger_payouts(transactions,code,'Principal Payout') if d < settlement)
                remaining=max(0,num(r.get('investment_amount'))-paid_before)
                principal_due.append({'due_date':settlement,'expected':remaining,'early_settlement':True,'original_due_date':final})
            else:
                principal_due=[{'due_date':settlement,'expected':max(0,num(r.get('investment_amount'))),'early_settlement':True,'original_due_date':final}]

        # Allocate principal cash sequentially; supports split transactions on the same day.
        pp=ledger_payouts(transactions,code,'Principal Payout'); pi=0
        for pd,pool0 in pp:
            pool=max(0,num(pool0))
            while pool>0.005 and pi<len(principal_due):
                e=principal_due[pi]; exp=max(0,num(e.get('expected')))
                actual=pool if e.get('early_settlement') or pi==len(principal_due)-1 else min(pool,exp if exp>0.005 else pool)
                e.update({'paid':True,'paid_date':pd,'actual':actual}); pool-=actual; pi+=1
        for e in principal_due:
            if 'paid' not in e: e.update({'paid':False,'paid_date':None,'actual':0})

        for typ,due_list in [('Profit',profit_due),('Principal',principal_due)]:
            for idx,e in enumerate(due_list):
                due=e['due_date']; expected=e['expected']; early=bool(e.get('early_settlement'))
                final_principal=(typ=='Principal' and (len(due_list)==1 or early))
                key=('principal-final' if final_principal else f'{typ.lower()}-{due.strftime("%Y-%m")}')
                if early: key=f'{typ.lower()}-early-{due.isoformat()}'
                if early and typ=='Profit':
                    label=f'Early settlement profit ({e.get("prorata_days",0)} days pro-rated)'
                elif early and typ=='Principal': label='Early principal settlement'
                else: label=('Final principal repayment' if final_principal else f'{due.strftime("%b %Y")} scheduled {typ.lower()}')
                base={'loan_code':code,'event_key':key,'event_type':typ,'label':label,'expected_date':due.isoformat(),
                      'expected_amount':expected,'is_paid':1 if e.get('paid') else 0,'paid_date':e['paid_date'].isoformat() if e.get('paid_date') else '',
                      'paid_amount':e.get('actual',0),'notes':'Auto-matched from transaction ledger' if e.get('paid') else '',
                      'early_settlement':early,'original_due_date':e.get('original_due_date').isoformat() if e.get('original_due_date') else '',
                      'prorata_days':e.get('prorata_days'),'prorata_from':e.get('prorata_from').isoformat() if e.get('prorata_from') else '',
                      'expected_gross':e.get('expected_gross')}
                manual=marks.get((code,key))
                if not manual and typ=='Profit' and not early:
                    manual=marks.get((code,'profit-'+due.strftime('%B').lower()))
                if manual and not e.get('paid'): base.update(manual)
                elif manual and e.get('paid') and manual.get('notes'): base['notes']=manual.get('notes')
                items.append(base)
        known={x['event_key'] for x in items}
        for (mc,mk),m in marks.items():
            if mc==code and mk not in known and not (mk.startswith('profit-') and any(x['event_type']=='Profit' and x.get('paid_date')==m.get('paid_date') for x in items)):
                items.append(m)
        items.sort(key=lambda x:(x.get('expected_date') or '9999-99-99',x.get('event_type') or ''))
        out[code]=items
    return out

def build_cash_projection(current_cash, last_tx_date, payment_schedule):
    events=[]
    for code,items in payment_schedule.items():
        for x in items:
            if x.get('event_type')=='Principal' and x.get('expected_date') and not x.get('is_paid'):
                events.append({'date':x['expected_date'],'type':'Receivable','category':'Principal','loan_code':code,'issuer':'','amount':num(x.get('expected_amount')),'notes':x.get('label','')})
    for x in cashflow_plan_rows():
        sign=-1 if x['flow_type']=='Allocation' else 1
        events.append({'date':x['flow_date'],'type':x['flow_type'],'category':x['flow_type'],'loan_code':x['loan_code'],'issuer':x['issuer'],'amount':sign*num(x['amount']),'notes':x['notes'],'id':x['id']})
    events=[x for x in events if x.get('date') and (not last_tx_date or x['date']>last_tx_date)]
    events.sort(key=lambda x:(x['date'],0 if x['amount']<0 else 1))
    bal=current_cash; timeline=[]
    for x in events:
        bal+=x['amount']; y=dict(x); y['projected_balance']=bal; timeline.append(y)
    return timeline

def build_payload():
    tx=load_transactions(); pf,schedule=load_simulation(); projection_history=load_projection_history();
    if not isinstance(projection_history, dict):
        projection_history={'cutoff':None,'events':projection_history if isinstance(projection_history,list) else []}
    mon=monitoring_map(); settings=get_settings(); issuer_limit=settings['issuer_limit']; pay_schedule=note_payment_schedule(pf,schedule,tx)
    profit_schedule=derived_profit_schedule(pf,tx,projection_history)
    for r in pf:
        r['monitoring']=mon.get(r['loan_code'], {'status':'Normal','owner':'','next_action':'','review_date':'','notes':''})
        r['payment_schedule']=pay_schedule.get(r['loan_code'],[])
    successful=[r for r in tx if str(r['status']).upper()=='SUCCESSFUL']
    latest_cash=successful[0]['current_balance'] if successful else (tx[0]['current_balance'] if tx else 0)
    invested=sum(r['amount'] for r in successful if r['action']=='Investment Committed')
    principal_paid=sum(r['amount'] for r in successful if r['action']=='Principal Payout')
    profit_paid=sum(r['amount'] for r in successful if r['action']=='Profit Payout')
    outstanding_principal=max(0, invested-principal_paid)
    deployed=outstanding_principal
    today_d=date.today(); q_start_month=((today_d.month-1)//3)*3+1; q_start=date(today_d.year,q_start_month,1)
    q_profit=sum(r['amount'] for r in successful if r['action']=='Profit Payout' and parse_date(r.get('date')) and q_start<=parse_date(r.get('date'))<=today_d)
    return_this_quarter=q_profit/10000000
    total_return=profit_paid/10000000
    active=[r for r in pf if is_active(r)]
    sim_outstanding=sum(max(0,r['unpaid_principal']) for r in active)
    unpaid_profit=sum(max(0,r['unpaid_profit']) for r in active)
    projected_net_profit=sum(max(0,r['net_profit']) for r in active)
    paid_profit_sim=sum(max(0,r['paid_profit']) for r in pf)
    service_fee=sum(max(0,r['service_fee']) for r in pf)
    total_value=latest_cash+deployed
    weighted_monthly=(sum(max(0,r['unpaid_principal'])*max(0,r['net_pm']) for r in active)/sim_outstanding) if sim_outstanding else 0
    today=date.today().isoformat()
    due=[]; daily={}
    for r in active:
        if r['final_date'] and re.match(r'^\d{4}-\d{2}-\d{2}$',r['final_date']):
            d=datetime.strptime(r['final_date'],'%Y-%m-%d').date(); days=(d-date.today()).days
            rr=dict(r); rr['days_to_due']=days; rr['expected_principal']=max(0,r['unpaid_principal']); due.append(rr)
            x=daily.setdefault(r['final_date'],{'date':r['final_date'],'principal':0,'notes':0,'items':[]})
            x['principal']+=max(0,r['unpaid_principal']); x['notes']+=1; x['items'].append({'loan_code':r['loan_code'],'issuer':r['issuer'],'amount':max(0,r['unpaid_principal']),'payment_type':r['payment_type']})
    due.sort(key=lambda x:x['days_to_due'])
    receivables=sorted(daily.values(),key=lambda x:x['date'])
    issuer={}
    for r in active:
        k=r['issuer'] or 'Unknown Issuer'; d=issuer.setdefault(k,{'issuer':k,'outstanding':0,'unpaid_profit':0,'notes':0,'rating':r['rating']})
        d['outstanding']+=max(0,r['unpaid_principal']); d['unpaid_profit']+=max(0,r['unpaid_profit']); d['notes']+=1
    issuers=[]
    for d in issuer.values():
        d['limit']=issuer_limit; d['headroom']=max(0,issuer_limit-d['outstanding']); d['over_limit']=max(0,d['outstanding']-issuer_limit); d['utilization']=d['outstanding']/issuer_limit if issuer_limit else 0
        d['allocatable_now']=min(latest_cash,d['headroom']); issuers.append(d)
    issuers.sort(key=lambda x:x['outstanding'],reverse=True)
    action_totals={}
    for r in successful: action_totals[r['action']]=action_totals.get(r['action'],0)+r['amount']
    balance_breakdown=[
        {'label':'Cash available','amount':latest_cash,'note':'Latest successful wallet balance'},
        {'label':'Total investment committed','amount':invested,'note':'Cumulative successful allocations since inception'},
        {'label':'Outstanding principal','amount':outstanding_principal,'note':'Investment committed less principal payouts'},
        {'label':'Unpaid net profit','amount':unpaid_profit,'note':'Expected future profit from active simulation notes'},
    ]
    last_tx_date=successful[0]['date'] if successful else (tx[0]['date'] if tx else None)
    try: sim_as_of=source_rows().get('simulation',{}).get('as_of','Unknown')
    except: sim_as_of='Unknown'
    projection=build_cash_projection(latest_cash,last_tx_date,pay_schedule)
    return {'settings':settings,'summary':{'cash':latest_cash,'ledger_deployed':deployed,'total_investment_committed':invested,'outstanding_principal':outstanding_principal,'ledger_profit_paid':profit_paid,'quarter_profit_collected':q_profit,'return_this_quarter':return_this_quarter,'total_return':total_return,'implied_fund_value':total_value,
                       'simulation_outstanding':sim_outstanding,'simulation_unpaid_profit':unpaid_profit,'active_notes':len(active),
                       'transaction_count':len(tx),'portfolio_count':len(pf),'last_transaction_date':last_tx_date,'simulation_as_of':sim_as_of,
                       'portfolio_gap':deployed-sim_outstanding,'projected_net_profit':projected_net_profit,'paid_profit_simulation':paid_profit_sim,
                       'weighted_monthly_net_rate':weighted_monthly,'annualized_net_rate':weighted_monthly*12,'service_fee':service_fee,
                       'investable_cash':latest_cash},
            'transactions':tx,'portfolio':pf,'schedule':schedule,'profit_schedule':profit_schedule,'due':due,'receivables':receivables,'issuers':issuers,
            'balance_breakdown':balance_breakdown,'action_totals':action_totals,'today':today,'cash_projection':projection,'cashflow_plan':cashflow_plan_rows(),'projection_history':{'cutoff':projection_history.get('cutoff'),'event_count':len(projection_history.get('events',[]))}}



@app.route('/api/data')
def data(): return jsonify(build_payload())

@app.route('/api/settings', methods=['POST'])
def save_settings():
    d=request.get_json(force=True) or {}; limit=num(d.get('issuer_limit'))
    if limit <= 0: return jsonify({'error':'Issuer limit must be greater than zero'}),400
    con=connect(); con.execute("INSERT INTO settings(key,value) VALUES('issuer_limit',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",(str(limit),)); con.commit(); con.close()
    return jsonify({'ok':True})

@app.route('/api/monitoring/<loan_code>', methods=['POST'])
def save_monitoring(loan_code):
    d=request.get_json(force=True) or {}
    vals=(loan_code,d.get('status','Normal'),d.get('owner',''),d.get('next_action',''),d.get('review_date',''),d.get('notes',''))
    con=connect()
    con.execute('''INSERT INTO monitoring(loan_code,status,owner,next_action,review_date,notes,updated_at)
      VALUES(?,?,?,?,?,?,CURRENT_TIMESTAMP) ON CONFLICT(loan_code) DO UPDATE SET status=excluded.status, owner=excluded.owner,
      next_action=excluded.next_action, review_date=excluded.review_date, notes=excluded.notes, updated_at=CURRENT_TIMESTAMP''',vals)
    con.commit(); con.close(); return jsonify({'ok':True})

@app.route('/api/payment-mark', methods=['POST'])
def save_payment_mark():
    d=request.get_json(force=True) or {}
    code=str(d.get('loan_code') or '').strip(); key=str(d.get('event_key') or '').strip()
    if not code or not key: return jsonify({'error':'loan_code and event_key are required'}),400
    vals=(code,key,d.get('event_type','Other'),d.get('label',''),d.get('expected_date',''),num(d.get('expected_amount')),1 if d.get('is_paid') else 0,d.get('paid_date',''),num(d.get('paid_amount')),d.get('notes',''))
    con=connect()
    con.execute("""INSERT INTO payment_marks(loan_code,event_key,event_type,label,expected_date,expected_amount,is_paid,paid_date,paid_amount,notes,updated_at)
      VALUES(?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP) ON CONFLICT(loan_code,event_key) DO UPDATE SET event_type=excluded.event_type,label=excluded.label,expected_date=excluded.expected_date,expected_amount=excluded.expected_amount,is_paid=excluded.is_paid,paid_date=excluded.paid_date,paid_amount=excluded.paid_amount,notes=excluded.notes,updated_at=CURRENT_TIMESTAMP""",vals)
    con.commit(); con.close(); return jsonify({'ok':True})

@app.route('/api/cashflow-plan', methods=['POST'])
def add_cashflow_plan():
    d=request.get_json(force=True) or {}; typ=d.get('flow_type')
    if typ not in ('Allocation','Receivable'): return jsonify({'error':'Flow type must be Allocation or Receivable'}),400
    if not d.get('flow_date') or num(d.get('amount'))<=0: return jsonify({'error':'Date and positive amount are required'}),400
    con=connect(); cur=con.execute('INSERT INTO cashflow_plan(flow_date,flow_type,loan_code,issuer,amount,notes) VALUES(?,?,?,?,?,?)',(d['flow_date'],typ,d.get('loan_code',''),d.get('issuer',''),num(d['amount']),d.get('notes',''))); con.commit(); i=cur.lastrowid; con.close(); return jsonify({'ok':True,'id':i})

@app.route('/api/cashflow-plan/<int:item_id>', methods=['DELETE'])
def delete_cashflow_plan(item_id):
    con=connect(); con.execute('DELETE FROM cashflow_plan WHERE id=?',(item_id,)); con.commit(); con.close(); return jsonify({'ok':True})

@app.route('/api/upload/<kind>', methods=['POST'])
def upload(kind):
    from uploads import validate_workbook
    if kind not in ('transactions','simulation','projections'):
        return jsonify({'error':'Invalid file type'}),400
    f=request.files.get('file')
    if not f or not f.filename.lower().endswith('.xlsx'):
        return jsonify({'error':'Please upload an .xlsx file'}),400
    data=f.read()
    as_of=request.form.get('as_of') or date.today().isoformat()
    if not parse_date(as_of):
        return jsonify({'error':'As-of date must be YYYY-MM-DD'}),400
    try:
        validate_workbook(kind,data)
        # Parse and calculate using only request-local candidate bytes. Never
        # replace the live workbook until all validation has succeeded.
        source_rows()
        g.sources[kind]={'kind':kind,'object_key':'candidate','as_of':as_of}
        if 'source_bytes' not in g: g.source_bytes={}
        g.source_bytes[kind]=data
        build_payload()
    except Exception:
        return jsonify({'error':'Workbook could not be read. Check its template, dates and amounts.'}),400
    finally:
        g.pop('sources',None)
        g.pop('source_bytes',None)
    backup=save_source(kind,data,as_of)
    return jsonify({'ok':True,'backup':backup or 'First import; no previous version'})


def simulation_update_snapshot(as_of=None):
    """Build copy/export values from the current simulation template, updated from successful ledger transactions.

    Only transaction-driven servicing fields are changed. Descriptive fields, remarks, note names and other template content
    remain exactly as they are in the imported simulation workbook.
    """
    cutoff=parse_date(as_of) if as_of else None
    tx=[t for t in load_transactions() if str(t.get('status','')).upper()=='SUCCESSFUL' and (not cutoff or (parse_date(t.get('date')) and parse_date(t.get('date'))<=cutoff))]
    portfolio,_=load_simulation(); by_code={r['loan_code']:r for r in portfolio}
    updates={}
    for code,r in by_code.items():
        note_tx=[t for t in tx if t.get('note')==code]
        inv=max(0,num(r.get('investment_amount')))
        paid_principal=sum(num(t.get('amount')) for t in note_tx if t.get('action')=='Principal Payout')
        tawidh=sum(num(t.get('amount')) for t in note_tx if t.get('action')=='Profit Payout' and ('tawidh' in str(t.get('transaction_no') or '').lower() or "ta'widh" in str(t.get('transaction_no') or '').lower()))
        paid_profit=sum(num(t.get('amount')) for t in note_tx if t.get('action')=='Profit Payout' and not ('tawidh' in str(t.get('transaction_no') or '').lower() or "ta'widh" in str(t.get('transaction_no') or '').lower()))
        unpaid_principal=max(0,inv-paid_principal)
        settlement=None; running=0.0
        principal_events=[]
        for t in note_tx:
            if t.get('action')!='Principal Payout': continue
            d=parse_date(t.get('date'))
            if d: principal_events.append((d,num(t.get('amount'))))
        principal_events.sort(key=lambda z:z[0])
        tol=max(1.0,inv*0.00001)
        for d,a in principal_events:
            running+=a
            if inv and running>=inv-tol:
                settlement=d; break
        final=parse_date(r.get('final_date'))
        completed=bool(inv and paid_principal>=inv-tol)
        early=bool(completed and settlement and final and settlement<final)
        original_net=max(0,num(r.get('net_profit')))
        original_gross=max(0,num(r.get('gross_profit')))
        if completed:
            expected_net=paid_profit
            unpaid_profit=0.0
        else:
            expected_net=original_net
            unpaid_profit=max(0,expected_net-paid_profit)
        gross_profit=(expected_net/0.8 if completed and expected_net else original_gross)
        service_fee=max(0,gross_profit-expected_net) if completed else max(0,num(r.get('service_fee')))
        updates[code]={
            'Loan Status':'Completed' if completed else r.get('loan_status',''),
            'Actual Repayment **':paid_principal+paid_profit,
            'Paid Principal':paid_principal,
            'Unpaid Principal':unpaid_principal,
            'Gross Profit':gross_profit,
            'Net Profit':expected_net,
            'Paid Profit':paid_profit,
            'Unpaid Profit':unpaid_profit,
            'Late Profit':tawidh,
            'Service Fee ':service_fee,
            'Early Repayment':'Early Repayment of Principal' if early else (r.get('early_repayment','') if not completed else ''),
            'Early Repayment Date':settlement.isoformat() if early and settlement else (r.get('early_date') if not completed else None),
        }
    # Read the displayed values from the template so copy/paste keeps all 36 columns and existing remarks/text.
    wbv=openpyxl.load_workbook(read_source('simulation'),data_only=True,read_only=True)
    wsv=wbv['Query result'] if 'Query result' in wbv.sheetnames else wbv[wbv.sheetnames[0]]
    headers=[c.value for c in next(wsv.iter_rows(min_row=1,max_row=1))]
    while headers and headers[-1] is None: headers.pop()
    rows=[]; template_codes=set()
    for vals in wsv.iter_rows(min_row=2,values_only=True):
        if not vals or vals[0] is None or not vals[1]: continue
        arr=list(vals[:len(headers)]); code=normalize_note_code(arr[1]); template_codes.add(code)
        up=updates.get(code,{})
        for i,h in enumerate(headers):
            if h in up:
                v=up[h]
                # Browser copy uses ISO dates; Excel accepts these cleanly when pasted.
                arr[i]=v
            elif isinstance(arr[i],(datetime,date)):
                arr[i]=arr[i].strftime('%Y-%m-%d')
        rows.append(arr)
    ledger_alloc={}
    for t in tx:
        if t.get('action')=='Investment Committed' and t.get('note'):
            ledger_alloc[t['note']]=ledger_alloc.get(t['note'],0)+num(t.get('amount'))
    missing=[{'loan_code':c,'allocated':a} for c,a in sorted(ledger_alloc.items()) if c not in template_codes]
    return {'headers':headers,'rows':rows,'updates':updates,'missing_notes':missing,'as_of':as_of or (max([t.get('date') for t in tx if t.get('date')],default=''))}


def build_updated_simulation_workbook(as_of=None):
    snap=simulation_update_snapshot(as_of)
    wb=openpyxl.load_workbook(read_source('simulation'),data_only=False)
    ws=wb['Query result'] if 'Query result' in wb.sheetnames else wb[wb.sheetnames[0]]
    headers=[c.value for c in ws[1]]
    while headers and headers[-1] is None: headers.pop()
    hidx={h:i+1 for i,h in enumerate(headers)}
    for row in range(2,ws.max_row+1):
        code=normalize_note_code(ws.cell(row,2).value)
        up=snap['updates'].get(code)
        if not up: continue
        # Inputs sourced from the ledger. Derived cells retain the template's formula-driven style.
        for h in ['Loan Status','Paid Principal','Paid Profit','Late Profit','Early Repayment','Early Repayment Date']:
            if h not in hidx: continue
            v=up.get(h)
            if h=='Early Repayment Date' and v:
                v=parse_date(v)
            ws.cell(row,hidx[h]).value=v
        # When a note is completed early, its expected profit itself changes to the realised pro-rated amount.
        if up.get('Early Repayment') and up.get('Loan Status')=='Completed':
            if 'Gross Profit' in hidx: ws.cell(row,hidx['Gross Profit']).value=up['Gross Profit']
            if 'Net Profit' in hidx: ws.cell(row,hidx['Net Profit']).value=up['Net Profit']
        # Re-apply the same core formulas used in the original MISB template for the transaction-driven fields.
        if 'Actual Repayment **' in hidx and 'Paid Principal' in hidx and 'Paid Profit' in hidx:
            ws.cell(row,hidx['Actual Repayment **']).value=f'={openpyxl.utils.get_column_letter(hidx["Paid Principal"])}{row}+{openpyxl.utils.get_column_letter(hidx["Paid Profit"])}{row}'
        if 'Unpaid Principal' in hidx and 'Investment Amount' in hidx and 'Paid Principal' in hidx:
            ws.cell(row,hidx['Unpaid Principal']).value=f'={openpyxl.utils.get_column_letter(hidx["Investment Amount"])}{row}-{openpyxl.utils.get_column_letter(hidx["Paid Principal"])}{row}'
        if 'Unpaid Profit' in hidx and 'Net Profit' in hidx and 'Paid Profit' in hidx:
            ws.cell(row,hidx['Unpaid Profit']).value=f'={openpyxl.utils.get_column_letter(hidx["Net Profit"])}{row}-{openpyxl.utils.get_column_letter(hidx["Paid Profit"])}{row}'
        if 'Service Fee ' in hidx and 'Gross Profit' in hidx and 'Net Profit' in hidx:
            ws.cell(row,hidx['Service Fee ']).value=f'={openpyxl.utils.get_column_letter(hidx["Gross Profit"])}{row}-{openpyxl.utils.get_column_letter(hidx["Net Profit"])}{row}'
    # Keep every sheet, comment/note, style, column width and existing Remarks cell from the imported template.
    try:
        wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode='auto'
    except Exception:
        pass
    return wb,snap

@app.route('/api/simulation-preview')
def simulation_preview():
    as_of=request.args.get('as_of') or None
    return jsonify(simulation_update_snapshot(as_of))

@app.route('/api/export/updated-simulation.xlsx')
def export_updated_simulation():
    as_of=request.args.get('as_of') or None
    wb,snap=build_updated_simulation_workbook(as_of)
    b=io.BytesIO(); wb.save(b); b.seek(0)
    stamp=(as_of or snap.get('as_of') or date.today().isoformat())
    return send_file(b,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',as_attachment=True,download_name=f'Simulation_Report_MISB_Updated_{stamp}.xlsx')

def report_workbook(p, as_of):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title='Fund Summary'
    green='0C7A58'; dark='17322B'; light='E8F4EF'; line='D9E4E0'
    ws['A1']='MISB Bi-Weekly Fund Monitoring Report'; ws['A1'].font=openpyxl.styles.Font(size=18,bold=True,color='FFFFFF'); ws['A1'].fill=openpyxl.styles.PatternFill('solid',fgColor=dark)
    ws.merge_cells('A1:D1'); ws['A2']='Report date'; ws['B2']=as_of; ws['A3']='Simulation source'; ws['B3']=p['summary']['simulation_as_of']; ws['A4']='Transaction ledger through'; ws['B4']=p['summary']['last_transaction_date']
    metrics=[('Cash available',p['summary']['cash']),('Total investment committed',p['summary']['total_investment_committed']),('Outstanding principal',p['summary']['outstanding_principal']),('Cumulative profit received',p['summary']['ledger_profit_paid']),('Return this quarter',p['summary']['return_this_quarter']),('Total returns since inception',p['summary']['total_return']),('Simulation outstanding principal',p['summary']['simulation_outstanding']),('Unpaid net profit',p['summary']['simulation_unpaid_profit']),('Issuer limit',p['settings']['issuer_limit'])]
    ws.append([]); ws.append(['Metric','Value'])
    for a,b in metrics: ws.append([a,b])
    for c in ws[6]: c.font=openpyxl.styles.Font(bold=True,color='FFFFFF'); c.fill=openpyxl.styles.PatternFill('solid',fgColor=green)
    for row in range(7,7+len(metrics)):
        ws.cell(row,2).number_format='0.00%' if 'return' in str(ws.cell(row,1).value).lower() else '"RM" #,##0.00'
    ws.column_dimensions['A'].width=36; ws.column_dimensions['B'].width=22; ws.column_dimensions['C'].width=18; ws.column_dimensions['D'].width=18

    sheets=[
      ('Issuer Exposure',['Issuer','Outstanding Principal','Limit','Remaining Allocation','Utilisation','Active Notes','Unpaid Profit'],
       [[x['issuer'],x['outstanding'],x['limit'],x['headroom'],x['utilization'],x['notes'],x['unpaid_profit']] for x in p['issuers']]),
      ('Receivables Calendar',['Date','Expected Principal','Notes Due'],[[x['date'],x['principal'],x['notes']] for x in p['receivables']]),
      ('Active Notes',['Loan Code','Issuer','Note Name','Status','Rating','Investment Amount','Outstanding Principal','Unpaid Profit','Final Repayment','Payment Type','Net p.m.','Monitoring Status','Owner','Next Action','Review Date'],
       [[r['loan_code'],r['issuer'],r['note_name'],r['loan_status'],r['rating'],r['investment_amount'],r['unpaid_principal'],r['unpaid_profit'],r['final_date'],r['payment_type'],r['net_pm'],r['monitoring'].get('status',''),r['monitoring'].get('owner',''),r['monitoring'].get('next_action',''),r['monitoring'].get('review_date','')] for r in p['portfolio'] if is_active(r)]),
      ('Transactions',['Date','Action','Note','Amount','Previous Balance','Current Balance','Transaction No','Pay Via'],
       [[r['date'],r['action'],r['note'],r['amount'],r['previous_balance'],r['current_balance'],r['transaction_no'],r['pay_via']] for r in p['transactions']])]
    for name,headers,rows in sheets:
        sh=wb.create_sheet(name); sh.append(headers)
        for c in sh[1]: c.font=openpyxl.styles.Font(bold=True,color='FFFFFF'); c.fill=openpyxl.styles.PatternFill('solid',fgColor=dark)
        for row in rows: sh.append(row)
        sh.freeze_panes='A2'; sh.auto_filter.ref=sh.dimensions
        for col in range(1,len(headers)+1): sh.column_dimensions[openpyxl.utils.get_column_letter(col)].width=min(36,max(12,len(headers[col-1])+3))
        for row in sh.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value,(int,float)): c.number_format='#,##0.00'
    return wb

@app.route('/api/export/biweekly.xlsx')
def export_biweekly_xlsx():
    p=build_payload(); as_of=request.args.get('as_of') or date.today().isoformat(); wb=report_workbook(p,as_of)
    b=io.BytesIO(); wb.save(b); b.seek(0)
    return send_file(b,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',as_attachment=True,download_name=f'MISB_Biweekly_Report_{as_of}.xlsx')


@app.route('/api/export/account-statement.xlsx')
def export_account_statement():
    month=request.args.get('month') or date.today().strftime('%Y-%m')
    try:
        start=datetime.strptime(month+'-01','%Y-%m-%d').date(); end=add_months(start,1)-timedelta(days=1)
    except:
        return jsonify({'error':'Month must be YYYY-MM'}),400
    rows=[]
    for r in load_transactions():
        d=parse_date(r.get('date'))
        if d and start<=d<=end and str(r.get('status','')).upper()=='SUCCESSFUL': rows.append(r)
    rows.sort(key=lambda r:(r.get('date') or '',r.get('id') or 0))
    wb=openpyxl.Workbook(); ws=wb.active; ws.title='Account Statement'
    dark='17322B'; green='0C7A58'
    ws['A1']='MISB Monthly Account Statement'; ws.merge_cells('A1:H1'); ws['A1'].font=openpyxl.styles.Font(size=18,bold=True,color='FFFFFF'); ws['A1'].fill=openpyxl.styles.PatternFill('solid',fgColor=dark)
    ws['A2']='Statement month'; ws['B2']=month
    opening=rows[0]['previous_balance'] if rows else 0; closing=rows[-1]['current_balance'] if rows else opening
    # Use the ledger's balance movement itself, so every action type is captured without assumptions.
    movements=[num(r.get('current_balance'))-num(r.get('previous_balance')) for r in rows]
    cash_in=sum(x for x in movements if x>0)
    cash_out=sum(-x for x in movements if x<0)
    ws['A4']='Opening balance'; ws['B4']=opening; ws['A5']='Cash in'; ws['B5']=cash_in; ws['A6']='Cash out'; ws['B6']=cash_out; ws['A7']='Closing balance'; ws['B7']=closing
    for rr in range(4,8): ws.cell(rr,2).number_format='"RM" #,##0.00'
    headers=['Date','Action','Note','Transaction No','Amount (RM)','Previous Balance (RM)','Current Balance (RM)','Pay Via']
    ws.append([]); ws.append(headers)
    hr=ws.max_row
    for c in ws[hr]: c.font=openpyxl.styles.Font(bold=True,color='FFFFFF'); c.fill=openpyxl.styles.PatternFill('solid',fgColor=green)
    for r in rows:
        ws.append([r['date'],r['action'],r['note'],r['transaction_no'],r['amount'],r['previous_balance'],r['current_balance'],r['pay_via']])
    for row in ws.iter_rows(min_row=hr+1):
        for c in row[4:7]: c.number_format='"RM" #,##0.00'
    widths=[14,24,14,36,18,22,22,14]
    for i,w in enumerate(widths,1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    ws.freeze_panes=f'A{hr+1}'; ws.auto_filter.ref=f'A{hr}:H{ws.max_row}'
    b=io.BytesIO(); wb.save(b); b.seek(0)
    return send_file(b,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',as_attachment=True,download_name=f'MISB_Account_Statement_{month}.xlsx')

@app.route('/api/export/biweekly.csv')
def export_biweekly():
    p=build_payload(); output=io.StringIO(); w=csv.writer(output); s=p['summary']
    w.writerow(['MISB Bi-Weekly Monitoring Snapshot', datetime.now().strftime('%Y-%m-%d')]); w.writerow([]); w.writerow(['Metric','Amount / Count'])
    for k in ['cash','ledger_deployed','implied_fund_value','ledger_profit_paid','simulation_outstanding','simulation_unpaid_profit','active_notes','portfolio_gap']:
        w.writerow([k,s[k]])
    w.writerow([]); w.writerow(['Issuer','Outstanding','Limit','Remaining Allocation','Utilisation'])
    for x in p['issuers']: w.writerow([x['issuer'],x['outstanding'],x['limit'],x['headroom'],x['utilization']])
    b=io.BytesIO(output.getvalue().encode('utf-8-sig')); b.seek(0)
    return send_file(b,mimetype='text/csv',as_attachment=True,download_name='MISB_Biweekly_Monitoring.csv')


from web import configure_web
configure_web(app)
from statements import register_statement_routes
register_statement_routes(app)
