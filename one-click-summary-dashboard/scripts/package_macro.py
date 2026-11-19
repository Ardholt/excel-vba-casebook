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


import argparse
p = argparse.ArgumentParser()
p.add_argument('--input', default='input')
p.add_argument('--output', default='output')
p.add_argument('--verify', action='store_true')
a = p.parse_args()
input_dir, output_dir = Path(a.input), Path(a.output)
output_dir.mkdir(parents=True, exist_ok=True)
source = Path('vba/SalesDashboard.bas').read_text()
if not a.verify:
    with tempfile.TemporaryDirectory(prefix='.macro-', dir=output_dir) as tmp:
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
        rows=read_sales(input_dir / ('before.xlsx' if (input_dir/'before.xlsx').exists() else 'input.xlsx'))
        accepted,audit,report=analyze(rows)
        make_book(output_dir/'sales-dashboard.xlsm',rows,accepted,audit,report,'BEFORE REFRESH',vba_project=binpath)

with ExcelFile(output_dir/'sales-dashboard.xlsm') as book:
    recovered=book.get_module('Module1')
    assert normalized(recovered) == normalized(source)
    names=book.module_names
    if callable(names): names=names()

package=importlib.metadata.distribution('pyOpenVBA')
licences=[p for p in package.files if p.name.lower() in ('license','license.txt','license.md')]
assert licences, 'The packaging template licence must accompany the workbook'
assert Path('vba/PACKAGING-LICENSE.txt').read_text() == package.locate_file(licences[0]).read_text()
with ZipFile(output_dir/'sales-dashboard.xlsm') as archive:
    binary = archive.read('xl/vbaProject.bin')
record_path = output_dir/'macro-package.json'
record = json.loads(record_path.read_text()) if a.verify else {}
if a.verify:
    record['packaged_input_sha256'] = record['workbook_sha256']
record.update({
    'workbook':'sales-dashboard.xlsm','module':'Module1',
    'source':'vba/SalesDashboard.bas','packaging_library':'pyOpenVBA '+package.version,
    'source_matches_embedded_module':True,'project_modules':list(names),
    'source_sha256':hashlib.sha256(source.encode()).hexdigest(),
    'vba_project_sha256':hashlib.sha256(binary).hexdigest(),
    'workbook_sha256':hashlib.sha256((output_dir/'sales-dashboard.xlsm').read_bytes()).hexdigest(),
    'source_verified_after_button_run_and_save':a.verify,
    'scope':'Embedded source extraction and identity; button execution recorded in vba-run.json'
})
record_path.write_text(json.dumps(record,indent=2)+'\n')
print('PASS: XLSM contains the readable module without source changes')
