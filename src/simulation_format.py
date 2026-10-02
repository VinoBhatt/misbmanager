"""September 2026 simulation: projections and realised returns remain separate."""
from datetime import date, datetime
from decimal import Decimal

ALIASES = {
    'Gross Profit': 'Gross Profit Projected',
    'Gross Profit Earned': 'Total Gross Profit Earned',
    'Net Profit': 'Total Net Profit Projected',
    'Paid Profit': 'Total Paid Out Profit',
    'Service Fee ': 'Total Service Fee Charged',
    'SST': 'SST Charged',
}
NEW_REQUIRED = {'Gross Profit Projected', 'Total Gross Profit Earned',
    'Total Net Profit Projected', 'Total Paid Out Profit', 'Total Service Fee Charged',
    'SST Charged', 'Net Late Payment Charges (SST Affected)',
    'Net Late Payment Charges (Non SST Affected)', 'SST Net Profit (After 1 Aug)',
    'Non SST Net Profit (Before 1 Aug)', 'SST Control Check'}


def canonical_row(row):
    if 'Gross Profit Projected' in row:
        result={**row, **{old:row.get(new) for old,new in ALIASES.items()}}
        early=bool(row.get('Early Repayment')) and float(row.get('Unpaid Principal') or 0)<=1
        result['Unpaid Profit']=0 if early else max(0,float(row.get('Total Net Profit Projected') or 0)-float(row.get('Total Paid Out Profit') or 0))
        return result
    return row


def column_positions(headers, fields):
    return {field:headers.index(field if field in headers else ALIASES[field])
            for field in fields if field in headers or ALIASES.get(field) in headers}


def realised_values(projected_gross, transactions, early=False):
    components = {name:Decimal(0) for name in (
        'Net Late Payment Charges (SST Affected)', 'Net Late Payment Charges (Non SST Affected)',
        'SST Net Profit (After 1 Aug)', 'Non SST Net Profit (Before 1 Aug)')}
    for transaction in transactions:
        day = transaction.get('date')
        if isinstance(day, datetime): day = day.date()
        if not isinstance(day, date): day = date.fromisoformat(str(day)[:10])
        taxable = day >= date(2026,8,1)
        reference = str(transaction.get('transaction_no') or '').lower().replace(chr(0x2019),"'")
        late = 'tawidh' in reference or "ta'widh" in reference
        key = ('Net Late Payment Charges (SST Affected)' if taxable else
               'Net Late Payment Charges (Non SST Affected)') if late else (
               'SST Net Profit (After 1 Aug)' if taxable else 'Non SST Net Profit (Before 1 Aug)')
        components[key] += Decimal(str(transaction.get('amount') or 0))
    taxable = components['Net Late Payment Charges (SST Affected)'] + components['SST Net Profit (After 1 Aug)']
    untaxed = components['Net Late Payment Charges (Non SST Affected)'] + components['Non SST Net Profit (Before 1 Aug)']
    actual_gross = taxable / Decimal('.784') + untaxed / Decimal('.8')
    fees = actual_gross * Decimal('.2')
    sst = taxable / Decimal('.784') * Decimal('.2') * Decimal('.08')
    projected = Decimal(str(projected_gross))
    net = projected * Decimal('.8')
    paid = taxable + untaxed
    return {**{key:float(value) for key,value in components.items()},
        '_new_logic':True, 'Gross Profit Projected':float(projected),
        'Total Gross Profit Earned':float(actual_gross), 'Total Service Fee Charged':float(fees),
        'SST Charged':float(sst), 'Total Net Profit Projected':float(net),
        'Total Paid Out Profit':float(paid), 'SST Control Check':.08 if taxable else 0,
        'Unpaid Profit':0 if early else float(max(Decimal(0),net-paid)),
        'Gross Profit':float(projected), 'Net Profit':float(net), 'Paid Profit':float(paid),
        'Service Fee ':float(fees), 'SST':float(sst)}
