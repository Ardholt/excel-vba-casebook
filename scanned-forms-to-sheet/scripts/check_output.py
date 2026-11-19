"""Check the produced workbooks against the field-level extraction."""
import csv,json,sys
from pathlib import Path
import openpyxl
from extract import FIELDS,data_rows
out=Path(sys.argv[1])
with (out/'fields.csv').open() as f: records=list(csv.DictReader(f))
report=json.loads((out/'report.json').read_text())
for name,raw in [('extracted.xlsx',True),('reviewed.xlsx',False)]:
    rows=list(openpyxl.load_workbook(out/name,data_only=True)['Data'].values)
    expected=data_rows(records,raw)
    assert len(rows)==len(expected)
    for actual,wanted in zip(rows,expected):
        assert tuple('' if x is None else x for x in actual)==tuple('' if x is None else x for x in wanted)
assert len(records)==report['input_forms']*len(FIELDS)
assert all(report['checks'].values())
print('PASS: workbook values and field conservation')
