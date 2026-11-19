"""Rebuild the worksheet pictures, animation and hero from the current results."""
import hashlib
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
report = json.loads(Path('output/report.json').read_text())
title = f"{report['duplicate_records']} repeated customer entries removed"
run('bash', str(Path(__file__).resolve().parents[2] / 'tools/hero/make-hero'),
    'images/before.png', 'images/after.png', 'media/hero.png',
    '--title', title,
    '--before-note', 'Same customer entry listed twice',
    '--after-note', 'Repeat removed from cleaned list',
    '--before-where', 'input.xlsx · Records',
    '--after-where', 'customer-register.xlsx · Records',
    '--tag', 'EXCEL & VBA CASEBOOK · CASE 01',
    '--before-crop', '780x247+0+141', '--after-crop', '780x247+0+141')
record = json.loads(Path('evidence/hero.json').read_text())
record.pop('temporary_renderer_assets_removed', None)
record['renderer'] = '../tools/hero/make-hero'
record['title'] = title
record['report_sha256'] = hashlib.sha256(Path('output/report.json').read_bytes()).hexdigest()
record['sha256'] = hashlib.sha256(Path('media/hero.png').read_bytes()).hexdigest()
record['sheet_render_seconds'] = render_seconds
record['total_seconds'] = round(time.perf_counter() - started, 4)
Path('evidence/hero.json').write_text(json.dumps(record, indent=2) + '\n')
