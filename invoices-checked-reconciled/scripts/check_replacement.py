"""Exercise moved headers, literal IDs, separate paths and two complete repeat runs."""
from pathlib import Path
from tempfile import TemporaryDirectory
from decimal import Decimal
import hashlib
import json
import subprocess
import time
from openpyxl import Workbook
from reconcile import HEADERS, save_json

start = time.perf_counter()
with TemporaryDirectory(prefix='.replacement-', dir='output') as temp:
    base = Path(temp)
    source = base / 'input with spaces'
    target = base / 'output with spaces'
    source.mkdir()
    target.mkdir()
    keep = target / 'keep.txt'
    keep.write_text('Keep this unrelated output file.\n')
    book = Workbook()
    book.remove(book.active)
    rows = {'Lines': [['BILL*1','L1','1','2.675','0.2'], ['bill*1','L1','-1','10','0.05'], ['000045','L1','2','3.335','0.2']],
            'Invoices': [['BILL*1','3.22'],['bill*1','-10.50'],['000045','8.01']],
            'Payments': [['X','BILL*1','1.22'],['Y','bill*1','-5.50'],['Z','000045','8'],['Z','000045','8'],['P','missing','2'],['plus?','000045','3']]}
    for name, headers in HEADERS.items():
        sheet = book.create_sheet(name)
        sheet.append(['Input for checking'])
        sheet.append([])
        sheet.append(list(reversed(headers)))
        for row in rows[name]:
            sheet.append(list(reversed(row)))
    filename = source / 'input.xlsx'
    book.save(filename)
    before = hashlib.sha256(filename.read_bytes()).hexdigest()
    results = []
    times = []
    for repeat in range(2):
        started = time.perf_counter()
        subprocess.run(['bash','scripts/inner.sh','--input',str(source),'--output',str(target)], check=True)
        times.append(round(time.perf_counter() - started, 3))
        results.append({name: (target / name).read_bytes() for name in ('invoices.csv','lines.csv','payments.csv','exceptions.csv')})
    import csv
    ledger = {row['invoice_id']: row for row in csv.DictReader((target / 'invoices.csv').open())}
    assert {key: Decimal(row['balance']) for key, row in ledger.items()} == {'BILL*1': Decimal('2'), 'bill*1': Decimal('-5'), '000045': Decimal('5')}
    report = json.loads((target / 'report.json').read_text())
    assert report['payment_rows']['duplicate_id'] == 2
    assert report['payment_rows']['unallocated'] == 1
    assert report['declared_mismatches'] == 1
    assert results[0] == results[1]
    assert hashlib.sha256(filename.read_bytes()).hexdigest() == before
    assert keep.read_text() == 'Keep this unrelated output file.\n'
    save_json('evidence/replacement-input.json', {'result':'PASS', 'runs':2,
        'checks':['Moved and reversed headers', 'Directories with spaces', 'Case-sensitive IDs with wildcard characters', 'Leading-zero ID retained', 'Half-up positive and negative calculations', 'Repeated payment IDs held', 'Unmatched payment retained', 'Same CSV bytes on second run', 'Input bytes unchanged', 'Unrelated output file retained'],
        'expected_balances': {'BILL*1':'2.00','bill*1':'-5.00','000045':'5.00'},
        'observed_balances': {k:v['balance'] for k,v in ledger.items()},
        'run_seconds':times, 'total_seconds':round(time.perf_counter()-start,3),
        'software_record':'runtime.json', 'recipe':'../scripts/check_replacement.py'})
