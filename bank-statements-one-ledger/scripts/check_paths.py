"""Exercise replacement inputs and repeat CLI runs using directories with spaces."""
import csv
import hashlib
import json
import shutil
import subprocess
import time
from decimal import Decimal
from pathlib import Path

start=time.perf_counter()
base=Path('output/.path-check')
source=base/'my statements'
target=base/'checked ledger'
shutil.copytree('input',source,dirs_exist_ok=True)
with (source/'layout-a.csv').open(newline='') as f: rows=list(csv.reader(f))
headers=rows[0]
credit=headers.index('credit'); balance=headers.index('balance')
rows[1][credit]='500.00'
for row in rows[1:]: row[balance]=format(Decimal(row[balance])+Decimal('50.00'),'.2f')
with (source/'layout-a.csv').open('w',newline='') as f: csv.writer(f).writerows(rows)
settings=json.loads((source/'accounts.json').read_text())
settings['layout-a.csv']['closing_balance']='10858.28'
(source/'accounts.json').write_text(json.dumps(settings,indent=2)+'\n')

def hashes():
    return {p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(source.iterdir()) if p.is_file()}

before=hashes(); reports=[]; runs=[]
for _ in range(2):
    run_start=time.perf_counter()
    result=subprocess.run(['bash','scripts/inner.sh','--input',str(source),'--output',str(target)],capture_output=True,text=True)
    assert result.returncode==0, 'Replacement-input CLI run failed: '+result.stdout
    report=json.loads((target/'report.json').read_text())
    assert report['input_records']==report['output_records']==810
    assert report['total_credits']=='54594.50'
    assert report['accounts'][0]['calculated_close']=='10858.28'
    report.pop('runtime_seconds'); reports.append(report)
    runtime=json.loads((target/'runtime.json').read_text())
    runs.append({'seconds':round(time.perf_counter()-run_start,6),'result':runtime['result'],'tests':runtime['tests']})
assert reports[0]==reports[1]
assert before==hashes()
Path('evidence/input-options.json').write_text(json.dumps({'command':'bash scripts/inner.sh --input "my statements" --output "checked ledger"',
    'replacement':{'layout-a.csv:first_credit':'500.00','expected_total_credits':'54594.50','expected_account_a_close':'10858.28'},
    'checks':{'paths_with_spaces':True,'replacement_values_used':True,'sources_unchanged':True,'repeat_business_result_identical':True},
    'runs':runs,'runtime_seconds':round(time.perf_counter()-start,6),'result':'PASS'},indent=2)+'\n')
shutil.rmtree(base)
print('Replacement-input CLI and repeat-run checks: PASS')
