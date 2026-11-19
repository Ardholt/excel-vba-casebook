import json
from pathlib import Path
from decimal import Decimal
import sys
import openpyxl
import pytest
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from combine import combine, collect, REQUIRED


def make(path, headers, rows, sheet='Data'):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


def test_public_conservation():
    report = json.loads(Path('output/report.json').read_text())
    headers, rows, manifest = collect('input/files')
    assert report['source_workbooks'] == 50
    assert report['files'] == dict(accepted=46, empty=2, rejected=2, ignored=1)
    assert len(rows) == report['output_rows'] == 5520
    assert sum(r[2] for r in manifest) == len(rows)
    assert sum(Decimal(str(r['amount'])) for r in rows) == Decimal(report['amount_total'])
    assert len({(r['source_file'], r['source_row']) for r in rows}) == 5520


def test_public_values_and_extras():
    headers, rows, _ = collect('input/files')
    wb = openpyxl.load_workbook('output/consolidated.xlsx', data_only=True)
    assert list(wb['Combined'].values) == [tuple(headers)] + [tuple(r.get(h) for h in headers) for r in rows]
    assert {'branch','channel','batch_code'} <= set(headers)
    assert any(r.get('channel') is None for r in rows)
    assert any(r.get('channel') == 'Online' for r in rows)


def test_equal_looking_records_survive():
    _, rows, _ = collect('input/files')
    pairs = [r for r in rows if r['source_file'].startswith('01_') and r['source_row'] in (120,121)]
    assert len(pairs) == 2
    assert pairs[0]['amount'] == pairs[1]['amount'] == 42.5
    assert pairs[0]['record_id'] != pairs[1]['record_id']


def test_moved_headers_extra_fields_blank_rows_and_repeat(tmp_path):
    source = tmp_path/'source with spaces'
    source.mkdir()
    make(source/'a.xlsx', ['amount','category','record_id','date','zone'], [[12.34,'Parts','0007','2025-02-01','Z1'], [None]*5, [-2,'Parts','0007','2025-02-01','Z2']])
    make(source/'b.xlsx', REQUIRED, [['0007','2025-02-01','Parts',12.34]])
    a = combine(source,tmp_path/'result')
    b = combine(source,tmp_path/'result')
    assert a['output_rows'] == b['output_rows'] == 3
    assert a['amount_total'] == b['amount_total'] == '22.68'
    wb = openpyxl.load_workbook(tmp_path/'result/consolidated.xlsx',data_only=True)
    vals = list(wb['Combined'].values)
    assert vals[1] == ('0007','2025-02-01','Parts',12.34,'a.xlsx',2,'Z1')
    assert vals[2][5] == 4
    assert vals[3][-1] is None


def test_new_file_enters_refresh(tmp_path):
    source = tmp_path/'in'
    source.mkdir()
    make(source/'a.xlsx', REQUIRED, [['R1','2025-01-01','Parts',10]])
    assert combine(source,tmp_path/'out')['output_rows'] == 1
    make(source/'b.xlsx', REQUIRED, [['R1','2025-01-01','Parts',10]])
    report = combine(source,tmp_path/'out')
    assert report['output_rows'] == 2
    assert report['amount_total'] == '20.00'


@pytest.mark.parametrize('headers,sheet,reason', [
    (['record_id','date','category'],'Data','Missing required columns'),
    (REQUIRED,'Other','Missing Data sheet'),
    (REQUIRED+['amount'],'Data','Duplicate header'),
    (REQUIRED+['source_file'],'Data','Reserved provenance header'),
    (REQUIRED+['','tail'],'Data','Blank header')])
def test_invalid_workbook_is_visible(tmp_path,headers,sheet,reason):
    make(tmp_path/'bad.xlsx',headers,[],sheet)
    _, rows, manifest = collect(tmp_path)
    assert rows == []
    assert manifest == [['bad.xlsx','rejected',0,reason]]


def test_empty_corrupt_temporary_and_output(tmp_path):
    make(tmp_path/'empty.xlsx',REQUIRED,[])
    (tmp_path/'bad.xlsx').write_bytes(b'not a workbook')
    (tmp_path/'~$busy.xlsx').write_bytes(b'lock')
    make(tmp_path/'consolidated.xlsx',REQUIRED,[['R1','2025-01-01','Parts',999]])
    _, rows, manifest = collect(tmp_path)
    assert not rows
    assert {r[0]:r[1] for r in manifest} == {'empty.xlsx':'empty','bad.xlsx':'rejected','~$busy.xlsx':'ignored','consolidated.xlsx':'ignored'}


def test_input_unchanged():
    import hashlib
    report = json.loads(Path('output/report.json').read_text())
    for name,digest in report['source_sha256'].items():
        assert hashlib.sha256((Path('input/files')/name).read_bytes()).hexdigest() == digest


def test_output_cannot_be_inside_input(tmp_path):
    with pytest.raises(ValueError,match='separate'):
        combine(tmp_path,tmp_path/'output')


def test_actual_vba_cells_equal_reference():
    expected = openpyxl.load_workbook('output/consolidated.xlsx',data_only=True)
    actual = openpyxl.load_workbook('output/vba-result.xlsx',data_only=True)
    for sheet in ('Combined','Manifest'):
        assert list(actual[sheet].values) == list(expected[sheet].values)
    result = json.loads(Path('evidence/vba-run.json').read_text())
    assert result['embedded_source_matches']
    assert [r['rows'] for r in result['runs']] == [5520,5520,5521]
    assert all(r['cells_equal_reference'] for r in result['runs'])
