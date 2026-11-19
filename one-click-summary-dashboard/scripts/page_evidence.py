"""Record renderer versions and strip incidental PNG metadata after capture."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from PIL import Image

probe = '''import json,importlib.metadata
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    browser=p.chromium.launch(headless=True,args=['--no-sandbox'])
    print(json.dumps({'playwright':importlib.metadata.version('playwright'),'chromium':browser.version}))
    browser.close()
'''
# The renderer has an isolated interpreter; do not install packages into the reference environment.
versions = None
for executable in [Path(sys.executable), *sorted(Path('/opt').glob('*/bin/python'))]:
    result=subprocess.run([str(executable),'-c',probe],text=True,capture_output=True)
    if result.returncode == 0:
        versions=json.loads(result.stdout)
        break
if versions is None:
    raise RuntimeError('Cannot read the installed page renderer versions')
captures=[]
for path in sorted(Path('evidence/page-preview').glob('*.png')):
    with Image.open(path) as image:
        clean=Image.new('RGB',image.size,'white'); clean.paste(image.convert('RGB')); clean.save(path)
    captures.append({'file':str(path),'width':clean.width,'height':clean.height,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
Path('evidence/page-runtime.json').write_text(json.dumps({'renderer':'pagerender; GitHub-style local Markdown preview',
    **versions,'captures':captures,
    'scope':'Local README preview in Chromium, not a signed-out visit to the published GitHub or Pages site'},indent=2)+'\n')
