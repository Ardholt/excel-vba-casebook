"""Render the page offline while preserving its repository CI badge URL."""
import hashlib
import json
from pathlib import Path
import subprocess
import time

readme = Path('README.md')
original = readme.read_bytes()
url = 'https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg'
badge = f'[![Tests]({url})](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)' 
assert url in original.decode()
started = time.perf_counter()
try:
    readme.write_text(original.decode().replace(badge, ''))
    subprocess.run(['pagerender', 'README.md', 'evidence/page-preview'], check=True)
finally:
    readme.write_bytes(original)
Path('evidence/page-preview/preview-method.json').write_text(json.dumps({
    'readme_sha256': hashlib.sha256(original).hexdigest(),
    'renderer': 'pagerender; GitHub-style Markdown in Chromium',
    'ci_badge': 'External CI badge omitted from this network-isolated capture. No replacement badge is shown. The README retains its real CI image and workflow links.',
    'limitations': 'Repository CI status and live GitHub/Pages navigation are outside this offline capture.',
    'seconds': round(time.perf_counter() - started, 4)
}, indent=2) + '\n')
