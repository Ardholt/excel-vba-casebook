"""Rebuild the worksheet pictures, animation and hero from the current results."""
import json
import subprocess
import time
from pathlib import Path


def run(*command):
    subprocess.run(command, check=True)


started = time.perf_counter()
run('python3', 'scripts/build_proof.py')
render_seconds = {}
for name in ('before', 'after'):
    start = time.perf_counter()
    run('xlrender', f'output/proof-{name}.xlsx', f'output/render-{name}')
    render_seconds[name] = round(time.perf_counter() - start, 4)
run('python3', 'scripts/finish_images.py')
run('python3', 'scripts/build_hero.py')
record = json.loads(Path('evidence/hero.json').read_text())
record['sheet_render_seconds'] = render_seconds
record['total_seconds'] = round(time.perf_counter() - started, 4)
Path('evidence/hero.json').write_text(json.dumps(record, indent=2) + '\n')
