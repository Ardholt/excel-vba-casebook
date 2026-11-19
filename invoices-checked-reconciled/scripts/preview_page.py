"""Render the README offline with a plain results link in place of the remote CI image."""
import json
from pathlib import Path
import subprocess

source = Path('README.md')
preview = Path('README.preview.md')
badge = 'https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg'
text = source.read_text()
assert f']({badge})' in text
preview.write_text(text.replace(f'![Tests]({badge})', 'View test results', 1))
try:
    subprocess.run(['pagerender', str(preview), 'evidence/page'], check=True)
    manifest = Path('evidence/page/page-manifest.json')
    if manifest.exists():
        raw = manifest.read_text().replace(str(Path.cwd()) + '/', '').replace('README.preview.md', 'README.md')
        data = json.loads(raw)
        if isinstance(data, dict):
            data['offline_ci_badge'] = 'Remote CI image rendered as a plain results link for local capture; README keeps its original badge.'
        manifest.write_text(json.dumps(data, indent=2) + '\n')
finally:
    preview.unlink(missing_ok=True)
