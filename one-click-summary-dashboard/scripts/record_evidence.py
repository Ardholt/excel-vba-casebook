"""Refresh evidence pointers and verify the artifacts of the completed run."""
import hashlib
import json
import re
from datetime import datetime,timezone
from pathlib import Path
from PIL import Image
from zipfile import ZipFile
import xml.etree.ElementTree as ET

def digest(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,v):Path(p).write_text(json.dumps(v,indent=2)+'\n')
run=json.loads(Path('output/vba-run.json').read_text())
package=json.loads(Path('output/macro-package.json').read_text())
tests=json.loads(Path('output/tests.json').read_text())
assert run['result']==tests['result']=='PASS'
assert run['saved_workbook_reopened_and_button_retested']
assert digest(Path('output/sales-dashboard.xlsm'))==run['refreshed_workbook_sha256']==package['workbook_sha256']
for phase in run['runs']:assert digest(Path('output')/phase['saved_output'])==phase['saved_sha256']
for name in ['runtime.json','vba-run.json','macro-package.json']:
    Path('evidence',name).write_bytes(Path('output',name).read_bytes())
write('evidence/validation.json',{'actual_button_result':'PASS: baseline, append, repeat, and saved-workbook reopen',
    'reference':'output/runtime.json','macro_execution':'output/vba-run.json','macro_packaging':'output/macro-package.json',
    'tests':tests})
write('evidence/verification.json',{'command':'bash scripts/inner.sh','checked_utc':datetime.now(timezone.utc).isoformat(),
    'result':'PASS','test_summary':f"{tests['passed']} passed",'input_sha256':{str(p):digest(p) for p in Path('input').glob('*.xlsx')},
    'source_sha256':{str(p):digest(p) for pattern in ['scripts/*.py','scripts/*.sh','tests/*.py','vba/*.bas'] for p in Path('.').glob(pattern)},
    'scope':'Container execution of the same inner command used by run.sh; fresh button artifacts and Python business tests.'})
for p in Path('.').rglob('*.md'):
    for link in re.findall(r'\]\(([^)]+)\)',p.read_text()):
        if '://' in link or link.startswith('#'):continue
        target=link.split('#')[0]
        if target and not target.startswith('../'):
            assert (p.parent/target).exists(),(p,link)
workbooks=list(Path('input').glob('*.xlsx'))+list(Path('output').glob('*.xlsx'))+list(Path('output').glob('*.xlsm'))
for p in workbooks:
    with ZipFile(p) as z:
        for name in ['docProps/core.xml','docProps/app.xml']:
            tree=ET.fromstring(z.read(name))
            for e in tree.iter():
                if e.tag.rsplit('}',1)[-1] in ('creator','lastModifiedBy','Company','Manager'):
                    assert not e.text,(p,e.tag,e.text)
write('evidence/file-checks.json',{'workbooks_checked':len(workbooks),'identifying_workbook_properties':'empty',
    'local_case_links':'All resolve','current_artifact_hashes':'PASS',
    'scope':'Workbook creator/company properties, local links and current macro/snapshot hashes. Visual review is recorded separately.'})
print('PASS: current evidence, artifact hashes, workbook properties and local links')
