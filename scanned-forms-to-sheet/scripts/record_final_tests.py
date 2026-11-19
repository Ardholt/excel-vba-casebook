"""Record the business-test result after the page build."""
import subprocess,time,json,re
from pathlib import Path
start=time.perf_counter()
p=subprocess.run(['python3','-m','pytest','-q','-p','no:cacheprovider','tests'],text=True,capture_output=True)
print(p.stdout)
summary=[line for line in p.stdout.splitlines() if re.search(r'\d+ (passed|failed|error)',line)]
Path('evidence/final-tests.json').write_text(json.dumps({'command':'python3 -m pytest -q -p no:cacheprovider tests','summary':summary[-1] if summary else 'No summary','exit_code':p.returncode,'seconds':round(time.perf_counter()-start,3)},indent=2)+'\n')
raise SystemExit(p.returncode)
