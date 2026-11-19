import argparse
import json
import platform
import subprocess
import time
from pathlib import Path
import openpyxl
from metadata import clean_properties
from customers import read_customers, read_confirmed, resolve

parser = argparse.ArgumentParser()
parser.add_argument('--input', required=True)
parser.add_argument('--output', required=True)
args = parser.parse_args()
source, out = Path(args.input), Path(args.output)
started = time.perf_counter()
clean_properties(out/'checked.xlsx')
records = read_customers(source/'input.xlsx')
mapping, _, errors = resolve(records, read_confirmed(source/'confirmed.csv'))
assert not errors
book = openpyxl.load_workbook(out/'result.xlsx', data_only=True)
checked = openpyxl.load_workbook(out/'checked.xlsx', data_only=True)
for name in book.sheetnames:
    assert list(book[name].values) == list(checked[name].values), f'LibreOffice changed {name} values'
rows = list(book['Customers'].values)[1:]
audit = list(book['Audit'].values)[1:]
assert {r[0] for r in rows} == set(mapping.values())
assert {r[0]:r[1] for r in audit} == mapping
assert len(audit) == len(records)
assert all(r[1] in set(mapping.values()) for r in audit)
by_id = {r['customer_id']:r for r in records}
assert all(tuple(by_id[r[0]][h] or None for h in ['customer_id','name','postcode']) == tuple(r) for r in rows)
report = json.loads((out/'report.json').read_text())
assert len(rows) == report['retained_records']
assert len(records) == len(rows)+report['merged_records']
# No confirmation means no identity changes, independently of proposal scores.
baseline, _, _ = resolve(records, [])
assert all(k == v for k,v in baseline.items())
record = {'status':'PASS', 'environment':'Ubuntu 24.04 container', 'python':platform.python_version(), 'openpyxl':openpyxl.__version__, 'libreoffice':subprocess.check_output(['libreoffice','--version'], text=True).strip(), 'execution':'Python produces XLSX workbooks; LibreOffice headless reopens and exports checked.xlsx; openpyxl compares all sheet values.', 'verified_sheets':book.sheetnames, 'source_id_count':len(records), 'retained_id_count':len(rows), 'no_confirmation_unchanged_count':len(baseline), 'check_seconds':round(time.perf_counter()-started,4), 'limitations':['Proposals are limited to nonblank equal normalized area codes and name similarity at least 0.82; they may miss duplicates or suggest distinct customers.', 'Customer fields come from the retained target; original fields remain in Audit.']}
(out/'runtime.json').write_text(json.dumps(record, indent=2)+'\n')
print('PASS: LibreOffice round trip and source-to-result conservation')
