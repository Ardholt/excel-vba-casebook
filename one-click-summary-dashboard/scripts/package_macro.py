"""Package the readable module into the before-state workbook; no macro execution."""
from pathlib import Path
import hashlib
import importlib.metadata
import json
import tempfile
from zipfile import ZipFile
from pyopenvba import ExcelFile
from dashboard import read_sales, analyze, make_book


def normalized(text):
    return '\n'.join(line.strip() for line in text.splitlines() if line.strip() and not line.startswith('Attribute '))


source = Path('vba/SalesDashboard.bas').read_text()
with tempfile.TemporaryDirectory(prefix='.macro-', dir='output') as tmp:
    root = Path(tmp)
    template = root/'template.xlsm'
    project = ExcelFile.create_new(template)
    project.set_module('Module1', source)
    project.save()
    project.close()
    with ZipFile(template) as archive:
        binary = archive.read('xl/vbaProject.bin')
    binpath=root/'vbaProject.bin'
    binpath.write_bytes(binary)
    rows=read_sales('input/before.xlsx')
    accepted,audit,report=analyze(rows)
    make_book('output/sales-dashboard.xlsm',rows,accepted,audit,report,'BEFORE REFRESH',vba_project=binpath)

with ExcelFile('output/sales-dashboard.xlsm') as book:
    recovered=book.get_module('Module1')
    assert normalized(recovered) == normalized(source)
    names=book.module_names
    if callable(names): names=names()

package=importlib.metadata.distribution('pyOpenVBA')
licences=[p for p in package.files if p.name.lower() in ('license','license.txt','license.md')]
assert licences, 'The packaging template licence must accompany the workbook'
Path('vba/PACKAGING-LICENSE.txt').write_text(package.locate_file(licences[0]).read_text())
Path('evidence/macro-package.json').write_text(json.dumps({
    'workbook':'output/sales-dashboard.xlsm','module':'Module1',
    'source':'vba/SalesDashboard.bas','packaging_library':'pyOpenVBA '+package.version,
    'source_matches_embedded_module':True,'project_modules':list(names),
    'source_sha256':hashlib.sha256(source.encode()).hexdigest(),
    'vba_project_sha256':hashlib.sha256(binary).hexdigest(),
    'workbook_sha256':hashlib.sha256(Path('output/sales-dashboard.xlsm').read_bytes()).hexdigest(),
    'excel_execution':'not_run','scope':'Packaging and source extraction, not compilation or macro execution'
},indent=2)+'\n')
print('PASS: candidate XLSM contains the readable module without source changes')
