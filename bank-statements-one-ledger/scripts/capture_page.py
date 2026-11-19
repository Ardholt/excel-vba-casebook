"""Capture the README; disclose any unavailable remote CI badge in the preview."""
import hashlib
import json
import subprocess
import time
from pathlib import Path

start=time.perf_counter()
source=Path('README.md').read_text()
badge='https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg'
command=['pagerender','README.md','evidence/page-preview']
result=subprocess.run(command,capture_output=True,text=True)
omitted=False
if result.returncode:
    message=result.stdout+result.stderr
    if "Unresolved README images: ['"+badge+"']" not in message:
        raise RuntimeError('Page capture failed: '+message)
    omitted=True
    preview=Path('.page-render.md')
    badge_line='[![Tests]('+badge+')](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)'
    assert badge_line in source
    preview.write_text(source.replace(badge_line,''))
    try:
        result=subprocess.run(['pagerender',str(preview),'evidence/page-preview'],capture_output=True,text=True)
        if result.returncode: raise RuntimeError('Page capture failed: '+result.stdout+result.stderr)
    finally:
        preview.unlink(missing_ok=True)
assert Path('README.md').read_text()==source
Path('evidence/page-run.json').write_text(json.dumps({'command':'pagerender README.md evidence/page-preview',
    'runtime_seconds':round(time.perf_counter()-start,6),
    'readme_sha256':hashlib.sha256(source.encode()).hexdigest(),
    'remote_ci_badge_omitted_in_preview_only':omitted,
    'badge_url':badge,'preview_badge':'Remote badge image omitted; no replacement status shown',
    'readme_kept_unchanged':True,
    'captures':{str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(Path('evidence/page-preview').glob('*.png'))},
    'limitation':'GitHub-style README body; if the remote badge cannot load, only that badge below the fold is omitted.'},indent=2)+'\n')
print(result.stdout.strip())
print('Remote CI badge omitted in preview only:',omitted)
