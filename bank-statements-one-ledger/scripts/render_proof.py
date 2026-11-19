"""Capture proof sheets after the default run; execute inside the shared image."""
import hashlib
import json
import shutil
import subprocess
import time
from pathlib import Path
from PIL import Image
import PIL
from openpyxl import load_workbook
from metadata import clean_workbook

start=time.perf_counter()
images=Path('images'); images.mkdir(exist_ok=True)
evidence=Path('evidence'); evidence.mkdir(exist_ok=True)
records=[]
for workbook,sheet,name in [
    ('output/ledger.xlsx','Balance checks','balance-checks.png'),
    ('output/input-overview.xlsx','Statement layouts','statement-layouts.png'),
    ('output/ledger.xlsx','Trace a balance','trace-a-balance.png'),
    ('output/input-overview.xlsx','Source balance proof','hero-before.png'),
    ('output/ledger.xlsx','Closing balance proof','hero-after.png'),
]:
    destination=Path('output/.renders')/name.removesuffix('.png')
    destination.mkdir(parents=True,exist_ok=True)
    # Isolate the requested sheet: PDF export can include other sheets with print areas.
    book=load_workbook(workbook)
    for other in list(book.sheetnames):
        if other != sheet: book.remove(book[other])
    capture=destination/'capture.xlsx'
    book.save(capture); clean_workbook(capture)
    result=subprocess.run(['xlrender',str(capture),str(destination),'--sheet',sheet],capture_output=True,text=True)
    if result.returncode: raise RuntimeError('Sheet render failed: '+sheet)
    paths=list(destination.glob('*.png'))
    if len(paths)!=1: raise RuntimeError('Expected one sheet image')
    # Preserve rendered pixels; save a clean PNG without ancillary metadata.
    with Image.open(paths[0]) as rendered:
        clean=Image.new('RGB',rendered.size,'white'); clean.paste(rendered.convert('RGB'))
        clean.save(images/name,optimize=True)
        size=list(clean.size)
    records.append({'file':'images/'+name,'source':workbook,'sheet':sheet,'width':size[0],'height':size[1],
                    'source_sha256':hashlib.sha256(Path(workbook).read_bytes()).hexdigest(),
                    'sha256':hashlib.sha256((images/name).read_bytes()).hexdigest(),
                    'method':'Selected worksheet copied to a single-sheet workbook, LibreOffice export rasterized and trimmed; pixels preserved; metadata stripped'})
shutil.rmtree('output/.renders')
runtime=json.loads(Path('output/runtime.json').read_text())
count=runtime['tests']['passed']
(images/'tests.svg').write_text(f'''<svg xmlns="http://www.w3.org/2000/svg" width="238" height="22" role="img" aria-label="Container tests: {count} passed">
<rect width="132" height="22" fill="#183541"/><rect x="132" width="106" height="22" fill="#087e83"/>
<g fill="white" font-family="Verdana,sans-serif" font-size="11" text-anchor="middle"><text x="66" y="15">container tests</text><text x="185" y="15">{count} passed</text></g></svg>\n''')
Path('requirements.lock').write_text('# Captured Python package versions; shared Docker image remains the runtime.\nopenpyxl=='+runtime['software']['openpyxl']+'\npytest=='+runtime['software']['pytest']+'\nPillow=='+PIL.__version__+'\n')
(evidence/'runtime.json').write_text(Path('output/runtime.json').read_text())
(evidence/'images.json').write_text(json.dumps({'software':runtime['software'],'images':records,'runtime_seconds':round(time.perf_counter()-start,6)},indent=2)+'\n')
print('Proof sheets rendered and capture record written.')
