"""Render the README offline without the unavailable remote CI badge image."""
from pathlib import Path
import hashlib
import json
import subprocess

source=Path('README.md').read_text()
remote='https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg'
rendered=source.replace(f'![Tests]({remote})','Test workflow')
target=Path('.page-preview.md')
try:
    target.write_text(rendered)
    subprocess.run(['pagerender',str(target),'evidence/page-preview'],check=True)
finally:
    target.unlink(missing_ok=True)
Path('evidence/page-source.json').write_text(json.dumps({'readme_sha256':hashlib.sha256(source.encode()).hexdigest(),
    'offline_substitution':'The network-isolated capture omits only the remote CI badge image, retaining its workflow link as text; the README retains the real badge URL.',
    'ci_status_claimed_by_preview':False},indent=2)+'\n')
