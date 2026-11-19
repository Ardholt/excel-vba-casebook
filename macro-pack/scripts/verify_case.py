"""Run the public reference/test entry point and record the current deliverables."""
from pathlib import Path
import datetime, hashlib, json, re, subprocess, time


def hashes(paths):
    return {str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(set(paths))}

inputs=hashes(Path('input').glob('*'))
start=time.perf_counter()
result=subprocess.run(['bash','scripts/inner.sh'],text=True,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
Path('evidence/tests.txt').write_text(result.stdout)
print(result.stdout,end='')
result.check_returncode()
assert inputs==hashes(Path('input').glob('*')), 'Input files changed'
Path('evidence/reference-checks.txt').write_text('Python reference and output validation: PASS\nSee tests.txt for the complete current run.\n')
sources=[Path('run.sh'),*Path('scripts').glob('*.py'),*Path('scripts').glob('*.sh'),*Path('tests').glob('*.py'),*Path('vba').glob('*.bas')]
record={
    'command':'bash scripts/inner.sh',
    'checked_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
    'exit_code':result.returncode,
    'seconds':round(time.perf_counter()-start,3),
    'test_summaries':re.findall(r'^\d+ passed[^\n]*',result.stdout,re.M),
    'inputs_unchanged':True,
    'input_sha256':inputs,
    'scope':'Reference outputs and business tests executed in the container. Actual macro execution is recorded separately in runtime.json.',
    'source_sha256':hashes(sources),
}
Path('evidence/verification.json').write_text(json.dumps(record,indent=2)+'\n')
artifacts=[*Path('input').glob('*'),Path('macro-pack.xlsm'),*Path('output').rglob('*.xlsx'),*Path('output').glob('*.zip'),*Path('output').glob('*.csv'),*Path('vba').glob('*.bas'),Path('README.md'),Path('media/hero.png')]
Path('evidence/checksums.json').write_text(json.dumps(hashes(artifacts),indent=2)+'\n')
