import argparse
from pathlib import Path
import json
from decimal import Decimal
import openpyxl
from combine import collect
p = argparse.ArgumentParser()
p.add_argument('--input', default='input/files')
p.add_argument('--output', default='output')
a = p.parse_args()
headers, rows, manifest = collect(a.input)
wb = openpyxl.load_workbook(Path(a.output) / 'consolidated.xlsx', data_only=True)
actual = list(wb['Combined'].values)
assert actual[0] == tuple(headers)
assert actual[1:] == [tuple(r.get(h) for h in headers) for r in rows]
assert list(wb['Manifest'].values)[1:] == [tuple(None if v == '' else v for v in r) for r in manifest]
report = json.loads((Path(a.output) / 'report.json').read_text())
assert report['output_rows'] == len(rows)
assert all(report['checks'].values())
print(f"PASS: source-to-output comparison for {len(rows)} records")
