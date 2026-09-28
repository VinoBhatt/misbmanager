"""Ledger-backed register of allocation and disbursement email confirmations."""
import hashlib
import json
from datetime import date
from flask import jsonify, request, g
from storage import connect, audit_event


EMAIL_COLUMNS = {'Email on Allocation': 'allocation_email_date',
                 'Email on Disbursement': 'disbursement_email_date'}


def saved_tracking():
    con = connect()
    try:
        return {row['loan_code']: dict(row) for row in con.execute('SELECT * FROM allocation_email_tracking')}
    finally:
        con.close()


def tracking_register():
    from app import load_transactions, load_simulation, normalize_note_code, parse_date
    from allocations import allocation_map
    groups = {}
    unmatched = []
    for tx in load_transactions():
        if str(tx.get('status', '')).upper() != 'SUCCESSFUL' or tx.get('action') != 'Investment Committed':
            continue
        if not tx.get('note'):
            unmatched.append(tx)
        else:
            groups.setdefault(tx['note'], []).append(tx)
    load_simulation()
    displayed = getattr(g, 'simulation_display', {'headers': [], 'rows': []})
    simulation = {}
    for values in displayed['rows']:
        row = dict(zip(displayed['headers'], values))
        simulation[normalize_note_code(row.get('Loan Code'))] = row
    campaigns = allocation_map()
    saved = saved_tracking()
    rows = []
    def valid_date(value):
        parsed = parse_date(value)
        return parsed.isoformat() if parsed else ''
    for code, transactions in groups.items():
        source = simulation.get(code, {})
        campaign = campaigns.get(code, {})
        record = saved.get(code)
        events = sorted([str(t.get('id') or ''), str(t.get('transaction_no') or ''),
                         str(t.get('date') or ''), str(t.get('amount') or 0)] for t in transactions)
        signature = hashlib.sha256(json.dumps(events).encode()).hexdigest()
        dates = sorted({valid_date(t.get('date')) for t in transactions} - {''})
        allocation_email = record['allocation_email_date'] if record else valid_date(source.get('Email on Allocation'))
        disbursement_email = record['disbursement_email_date'] if record else valid_date(source.get('Email on Disbursement'))
        disbursal = valid_date(source.get('Disbursal Date')) or valid_date(campaign.get('disbursal_date'))
        changed = bool(record and record['ledger_signature'] != signature)
        # Imported dates cannot confirm later allocations that were not yet emailed.
        imported_review = bool(not record and dates and any(d and d < dates[-1] for d in (allocation_email, disbursement_email)))
        pending = []
        if not allocation_email: pending.append('Allocation email pending')
        if not disbursement_email: pending.append('Disbursement email pending')
        if changed or imported_review: pending.append('Review allocation coverage')
        if len(dates) == 0 or any(not valid_date(t.get('date')) for t in transactions): pending.append('Allocation date missing')
        if not disbursal: pending.append('Disbursement date missing')
        rows.append({'loan_code': code, 'issuer': source.get('Issuer Name') or campaign.get('issuer_name', ''),
                     'note_name': source.get('Note Name') or campaign.get('note_name', ''),
                     'amount': round(sum(float(t.get('amount') or 0) for t in transactions), 2),
                     'allocation_dates': dates, 'transaction_count': len(transactions),
                     'allocation_email_date': allocation_email, 'disbursement_email_date': disbursement_email,
                     'disbursal_date': disbursal, 'ledger_signature': signature,
                     'source': 'Saved confirmation' if record else 'Imported simulation' if source else 'Not in simulation',
                     'pending': pending, 'emails_complete': bool(allocation_email and disbursement_email and not changed and not imported_review)})
    rows.sort(key=lambda r: (not bool(r['pending']), r['loan_code']))
    return {'rows': rows, 'unmatched': unmatched,
            'summary': {'total': len(rows), 'emails_complete': sum(r['emails_complete'] for r in rows),
                        'pending': sum(not r['emails_complete'] for r in rows),
                        'missing_disbursal': sum(not r['disbursal_date'] for r in rows)}}


def apply_email_dates(headers, rows, updates):
    if 'Loan Code' not in headers: return
    from app import normalize_note_code
    saved = saved_tracking()
    for row in rows:
        code = normalize_note_code(row[headers.index('Loan Code')])
        record = saved.get(code)
        if record is None: continue
        for column, field in EMAIL_COLUMNS.items():
            if column in headers:
                row[headers.index(column)] = record[field] or 'Nil'
                updates.setdefault(code, {})[column] = record[field] or 'Nil'


def register_email_tracking_routes(app):
    @app.get('/api/allocation-email-tracking')
    def email_tracking():
        return jsonify(tracking_register())

    @app.put('/api/allocation-email-tracking/<code>')
    def save_email_tracking(code):
        from allocations import iso_date
        data = request.get_json(silent=True) or {}
        row = next((r for r in tracking_register()['rows'] if r['loan_code'] == code), None)
        if row is None: return jsonify(error='No successful ledger allocation found for this note.'), 404
        if data.get('ledger_signature') != row['ledger_signature']:
            return jsonify(error='The allocations changed. Refresh the monitor and review them before saving.'), 409
        try:
            dates = [iso_date(data.get(field), label, False) for field, label in
                     [('allocation_email_date', 'Allocation email date'), ('disbursement_email_date', 'Disbursement email date')]]
            if any(value > date.today().isoformat() for value in dates):
                raise ValueError('Sent dates cannot be in the future.')
        except ValueError as error:
            return jsonify(error=str(error)), 400
        con = connect()
        try:
            con.execute('''INSERT INTO allocation_email_tracking
                (loan_code,allocation_email_date,disbursement_email_date,ledger_signature) VALUES(?,?,?,?)
                ON CONFLICT(loan_code) DO UPDATE SET allocation_email_date=excluded.allocation_email_date,
                disbursement_email_date=excluded.disbursement_email_date,
                ledger_signature=excluded.ledger_signature,updated_at=CURRENT_TIMESTAMP''',
                (code, *dates, row['ledger_signature']))
            con.commit()
        finally:
            con.close()
        audit_event('Confirmed allocation emails', 'allocation', code,
                    dict(zip(('allocation_email_date', 'disbursement_email_date'), dates)))
        return jsonify(ok=True)
