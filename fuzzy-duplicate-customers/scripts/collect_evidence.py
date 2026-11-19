import hashlib
import json
from pathlib import Path
runtime=json.loads(Path('output/runtime.json').read_text())
runtime['business_report']='output/report.json'
runtime['test_results']='output/test-results.json'
runtime['inputs']={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in [Path('input/input.xlsx'),Path('input/confirmed.csv')]}
runtime['outputs']={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in [Path('output/result.xlsx'),Path('output/checked.xlsx'),Path('output/audit.csv'),Path('output/proposed.csv')]}
runtime['transformation_seconds']=json.loads(Path('output/report.json').read_text())['runtime_seconds']
runtime['tests']=json.loads(Path('output/test-results.json').read_text())
Path('evidence').mkdir(exist_ok=True)
Path('evidence/runtime.json').write_text(json.dumps(runtime,indent=2)+'\n')
