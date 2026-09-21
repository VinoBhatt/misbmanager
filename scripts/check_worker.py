"""Exercise local Cloudflare bindings; never connects to a remote deployment."""
import http.cookiejar
import io
import json
from pathlib import Path
import urllib.request
import urllib.error
import uuid
import openpyxl

ROOT=Path(__file__).resolve().parent.parent
config=dict(line.split('=',1) for line in (ROOT/'.dev.vars').read_text().splitlines() if '=' in line and not line.startswith('#'))
opener=urllib.request.build_opener(urllib.request.ProxyHandler({}),urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
base='http://localhost:8791'


def request(path, data=None, headers=None, method=None):
    req=urllib.request.Request(base+path,data=data,headers=headers or {},method=method)
    try:
        with opener.open(req,timeout=60) as response:
            return response.status,response.read()
    except urllib.error.HTTPError as error:
        return error.code,error.read()


def post(path,body):
    return request(path,json.dumps(body).encode(),{'Content-Type':'application/json'})


status,body=request('/login')
assert status==200,(status,body[:500])
assert b'Workspace password' in body
status,body=request('/api/data')
assert status==401,(status,body[:500])
status,body=post('/api/login',{'password':config['APP_PASSWORD']})
assert status==200,(status,body[:500])
status,body=request('/api/data')
assert status==200,(status,body[:500])
print('Worker login and D1 reads passed',flush=True)
for kind,filename in [('transactions','transactions.xlsx'),('simulation','simulation.xlsx'),('projections','midas_projections.xlsx')]:
    boundary=uuid.uuid4().hex
    data=(ROOT/'legacy/MISB_Fund_Tracker/data'/filename).read_bytes()
    payload=(f'--{boundary}\r\nContent-Disposition: form-data; name="as_of"\r\n\r\n2026-08-15\r\n'
             f'--{boundary}\r\nContent-Disposition: form-data; name="file"; filename="{filename}"\r\n'
             'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n').encode()+data+f'\r\n--{boundary}--\r\n'.encode()
    status,body=request('/api/upload/'+kind,payload,{'Content-Type':'multipart/form-data; boundary='+boundary})
    assert status==200,(kind,status,body[:500])
    print(f'Worker {kind} import and D1 workbook persistence passed',flush=True)
status,body=request('/api/data')
assert status==200,(status,body[:500])
data=json.loads(body)
assert data['summary']['portfolio_count']>0
status,body=post('/api/settings',{'issuer_limit':2000000})
assert status==200,(status,body[:500])
for export in ('biweekly.xlsx','account-statement.xlsx?month=2026-08','updated-simulation.xlsx?as_of=2026-08-15'):
    status,body=request('/api/export/'+export)
    assert status==200,(export,status,body[:500])
    wb=openpyxl.load_workbook(io.BytesIO(body));assert wb.sheetnames;wb.close()
    print('Worker export passed: '+export,flush=True)
print('Cloudflare runtime checks passed.',flush=True)
