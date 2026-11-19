"""Regression checks on artifacts produced by the actual button runner."""
import hashlib
import json
import os
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import openpyxl

ROOT=Path(__file__).resolve().parents[1]
OUT=Path(os.environ.get('DASHBOARD_RUN_OUTPUT',ROOT/'output'))

def record(name):return json.loads((OUT/name).read_text())
def digest(path):return hashlib.sha256(path.read_bytes()).hexdigest()

def test_saved_macro_and_source_match_this_run():
    package=record('macro-package.json');run=record('vba-run.json')
    assert package['source_verified_after_button_run_and_save']
    assert package['source_sha256']==run['source_sha256']==digest(ROOT/'vba/SalesDashboard.bas')
    assert package['workbook_sha256']==run['refreshed_workbook_sha256']==digest(OUT/'sales-dashboard.xlsm')

def test_download_contains_one_bound_refresh_button():
    with ZipFile(OUT/'sales-dashboard.xlsm') as archive:
        assert 'xl/vbaProject.bin' in archive.namelist()
        bindings=[]
        for name in archive.namelist():
            if name.endswith('.vml'):
                root=ET.fromstring(archive.read(name))
                bindings.extend(e.text for e in root.iter() if e.tag.endswith('}FmlaMacro'))
        assert bindings==['Module1.RefreshDashboard']

def test_actual_button_actions_and_saved_snapshots():
    run=record('vba-run.json')
    assert run['result']=='PASS' and run['saved_workbook_reopened_and_button_retested']
    assert run['button']['activation']=='XAccessibleAction.doAccessibleAction(0)'
    assert '.Module1.RefreshDashboard?' in run['button']['binding']
    assert [r['phase'] for r in run['runs']]==['before','after','repeat']
    for phase in run['runs']:
        assert phase['button_activated'] and phase['cleared_outputs_rebuilt']
        assert phase['all_accepted_rows_match'] and phase['all_audit_rows_match'] and phase['source_rows_unchanged']
        path=OUT/phase['saved_output']
        assert digest(path)==phase['saved_sha256']
        book=openpyxl.load_workbook(path,data_only=True)
        assert Decimal(str(book['Dashboard']['A3'].value))==Decimal(phase['paid_revenue'])
        book.close()

def test_delivered_workbook_is_the_refreshed_result():
    expected=record('report.json');run=record('vba-run.json')
    assert run['runs'][1]['paid_revenue']==run['runs'][2]['paid_revenue']==expected['paid_revenue']
    book=openpyxl.load_workbook(OUT/'sales-dashboard.xlsm',data_only=True)
    assert Decimal(str(book['Dashboard']['A3'].value))==Decimal(expected['paid_revenue'])
    assert book['Sales'].max_row-1==expected['input_records']
    book.close()

def test_vba_document_module_names_survive_export():
    ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with ZipFile(OUT/'sales-dashboard.xlsm') as z:
        root=ET.fromstring(z.read('xl/workbook.xml'))
        assert root.find('s:workbookPr',ns).get('codeName')=='ThisWorkbook'
        names=[(n.get('name'),n.get('localSheetId')) for n in root.findall('s:definedNames/s:definedName',ns)]
        assert len(names)==len(set(names))
        for i in range(1,5):
            root=ET.fromstring(z.read(f'xl/worksheets/sheet{i}.xml'))
            assert root.find('s:sheetPr',ns).get('codeName')==f'Sheet{i}'
