"""Bound spreadsheet expansion before handing XML to openpyxl."""
import io
import zipfile
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
