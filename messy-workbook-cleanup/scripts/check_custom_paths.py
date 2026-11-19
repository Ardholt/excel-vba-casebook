"""Exercise the documented input/output options using folders with spaces."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import time

started = time.perf_counter()
with tempfile.TemporaryDirectory(prefix='.checks-path ', dir='output') as directory:
    root = Path(directory)
    source, result = root / 'customer exports', root / 'cleaned results'
    shutil.copytree('input', source)
    subprocess.run(['bash', 'scripts/inner.sh', '--input', str(source), '--output', str(result)],
                   check=True, env={**os.environ, 'PYTHONDONTWRITEBYTECODE': '1'})
    report = json.loads((result / 'report.json').read_text())
    baseline = json.loads(Path('output/report.json').read_text())
    assert report['business_result_sha256'] == baseline['business_result_sha256']
    runtime = json.loads((result / 'runtime.json').read_text())
    assert runtime['checks']['source_files_unchanged']
    evidence = {
        'command': 'bash scripts/inner.sh --input <folder with spaces> --output <folder with spaces>',
        'checks': {'input_and_output_paths_with_spaces': True,
                   'business_result_matches_default': True, 'sources_unchanged': True},
        'business_result_sha256': report['business_result_sha256'],
        'software': runtime['software'], 'seconds': round(time.perf_counter() - started, 4),
        'test_result': (result / 'test-results.txt').read_text().strip(),
        'scope': 'Container command options; the outer Docker mount route is a repository-level check.'}
Path('evidence/path-options.json').write_text(json.dumps(evidence, indent=2) + '\n')
print('Input and output paths with spaces: PASS')
