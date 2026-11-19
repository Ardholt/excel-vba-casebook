"""Record the actual business-test outcome without environment identifiers."""
import argparse
import os
import json
import re
import subprocess
import sys
import time
from pathlib import Path

p = argparse.ArgumentParser()
p.add_argument('--output', default='output')
a = p.parse_args()
os.environ['DASHBOARD_RUN_OUTPUT']=str(Path(a.output).resolve())
started = time.perf_counter()
result = subprocess.run([sys.executable, '-m', 'pytest', '-q', '-p', 'no:cacheprovider', 'tests'], text=True, capture_output=True)
print(result.stdout, end='')
if result.stderr:
    print(result.stderr, end='', file=sys.stderr)
match = re.search(r'(\d+) passed', result.stdout)
record = {'result': 'PASS' if result.returncode == 0 and match else 'FAIL',
          'exit_code': result.returncode, 'passed': int(match.group(1)) if match else 0,
          'elapsed_seconds': round(time.perf_counter()-started,3),
          'command': 'python3 -m pytest -q -p no:cacheprovider tests', 'scope': 'Python business logic and current button-run artifact checks; button execution is recorded in vba-run.json'}
(Path(a.output)/'tests.json').write_text(json.dumps(record, indent=2)+'\n')
if record['result'] != 'PASS':
    raise SystemExit(result.returncode or 1)
