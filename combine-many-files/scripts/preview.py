"""Render the README offline, representing the remote CI badge as a text link."""
from pathlib import Path
import subprocess
import json
import tempfile

readme = Path('README.md').read_text()
badge = '[![Tests](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg)](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)'
linked_text = '[Tests: see status on GitHub](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)'
assert readme.count(badge) == 1
with tempfile.NamedTemporaryFile(mode='w', suffix='.md', prefix='.page-', dir='.', delete=False) as f:
    f.write(readme.replace(badge, linked_text))
    preview = Path(f.name)
try:
    subprocess.run(['pagerender', str(preview), 'images/page-preview'], check=True)
finally:
    preview.unlink(missing_ok=True)
manifest = Path('images/page-preview/page-manifest.json')
if manifest.exists():
    data = json.loads(manifest.read_text())
    # Retain measurements, replacing private or temporary path fields with public-relative paths.
    def relative(value):
        if isinstance(value, dict):
            return {k: relative(v) for k,v in value.items()}
        if isinstance(value, list):
            return [relative(v) for v in value]
        if isinstance(value, str):
            if str(preview) in value:
                return value.replace(str(preview),'README.md')
            if value.startswith('/work/'):
                return value[6:]
        return value
    data = relative(data)
    if isinstance(data, list):
        data = {'renders': data}
    data['offline_badge'] = 'Remote CI badge replaced by a link for offline rendering only. README keeps the required badge.'
    data['source'] = 'README.md'
    manifest.write_text(json.dumps(data, indent=2)+'\n')

import hashlib
import shutil
renderer = Path(shutil.which('pagerender'))
Path('evidence/page-render.json').write_text(json.dumps({
    'renderer': 'pagerender', 'renderer_sha256': hashlib.sha256(renderer.read_bytes()).hexdigest(),
    'viewports': [390, 1366], 'scope': 'Offline approximation of the GitHub README body; no GitHub navigation.',
    'remote_badge': 'Replaced with a linked text placeholder in preview only; README retains the CI badge.',
    'browser_version': 'Not recorded by the renderer'
}, indent=2) + '\n')
