"""Capture the page and record its source hashes and elapsed time."""
import hashlib
import json
import subprocess
import time
import sys
from pathlib import Path
from PIL import Image

start=time.perf_counter()
report=json.loads(Path('output/report.json').read_text())
count=report['output_records']
for stem in ('before','after'):
    subprocess.run(['convert',f'images/hero-{stem}.png','-resize','1280x','-background','white',
                    '-gravity','north','-extent','1280x406','-strip',f'media/{stem}.png'],check=True)
command=['bash',str(Path(__file__).resolve().parents[2] / 'tools/hero/make-hero'),'media/before.png','media/after.png','media/hero.png',
         '--title','One sheet; the 25p difference needs review',
         '--before-note','3 statements, balances unchecked',
         '--after-note',f'{count} kept; 25p still needs review',
         '--before-where','layout-c.csv', '--after-where','ledger.xlsx',
         '--tag','EXCEL & VBA CASEBOOK · CASE 02']
subprocess.run(command,check=True)
with Image.open('media/hero.png') as img: dimensions=list(img.size)
Path('evidence/hero.json').write_text(json.dumps({'source_before':'images/hero-before.png','source_after':'images/hero-after.png',
    'source_sha256':{p:hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in ['images/hero-before.png','images/hero-after.png','output/report.json']},
    'rendered_sha256':{p:hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in ['media/before.png','media/after.png','media/hero.png']},
    'dimensions':dimensions,'method':'Shared before/after style; exact worksheet values; panels enlarged and padded below.',
    'software':subprocess.check_output(['convert','-version'],text=True).splitlines()[0],
    'runtime_seconds':round(time.perf_counter()-start,6)},indent=2)+'\n')
subprocess.run([sys.executable,'scripts/capture_page.py'],check=True)
