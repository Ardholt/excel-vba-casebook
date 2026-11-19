"""Smoke-check separate input/output directories, numeric additions and zero amounts."""
import json
from pathlib import Path
import subprocess,tempfile
import openpyxl

with tempfile.TemporaryDirectory(prefix='dashboard-custom-') as tmp:
    root=Path(tmp);src=root/'input space';out=root/'output space';src.mkdir()
    rows=[['S1','2026-01-01','Parts',10.5,'Paid'],['S2','2026-02-01','Parts',0,'Paid']]
    for filename,values in [('before.xlsx',rows[:1]),('input.xlsx',rows)]:
        book=openpyxl.Workbook();sheet=book.active;sheet.title='Sales'
        sheet.append(['sale_id','date','category','amount','status'])
        for row in values:sheet.append(row)
        book.save(src/filename)
    subprocess.run(['bash','scripts/inner.sh','--input',str(src),'--output',str(out)],check=True)
    record=json.loads((out/'vba-run.json').read_text())
    assert record['result']=='PASS' and record['runs'][-1]['paid_revenue']=='10.50'
    assert record['runs'][-1]['paid_count']==2
    assert record['saved_workbook_reopened_and_button_retested']
    Path('evidence/custom-run.json').write_text(json.dumps({
        'result':'PASS','scope':'Separate directories with spaces; numeric amount and appended zero amount; full inner.sh workflow',
        'paid_revenue':record['runs'][-1]['paid_revenue'],'paid_count':2,
        'button_phases':[r['phase'] for r in record['runs']],
        'saved_workbook_reopened_and_button_retested':True},indent=2)+'\n')
print('PASS: custom input/output directories and numeric appended sales')
