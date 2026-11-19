import sys
from decimal import Decimal
from pathlib import Path
import pytest
import openpyxl
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from dashboard import analyze, build, read_sales
from generate_input import HEADERS, rows, ADDED


def records(data):
    return [{'source_row': i+2, **dict(zip(HEADERS,r))} for i,r in enumerate(data)]


def save(path, data, headers=HEADERS, offset=0):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = 'Sales'
    for i in range(offset):
        sheet.append(['Sales export'])
    sheet.append(headers)
    for r in data:
        d = dict(zip(HEADERS, r))
        sheet.append([d[h] for h in headers])
    book.save(path)


def test_reference_conserves_rows_and_money():
    valid, audit, report = analyze(records(rows()+ADDED))
    assert report['paid_revenue'] == '133102.36'
    assert (report['input_records'], report['paid_count'], report['pending_count'], report['cancelled_count'], report['rejected_records']) == (1332, 1149, 106, 45, 32)
    assert sum(Decimal(v) for v in report['paid_by_month'].values()) == Decimal('133102.36')
    assert sum(Decimal(v) for v in report['paid_by_category'].values()) == Decimal('133102.36')
    assert len(audit) == len(valid) + 32


def test_new_sales_create_month_category_and_include_return_once():
    before = analyze(records(rows()))[2]
    after = analyze(records(rows()+ADDED))[2]
    assert Decimal(after['paid_revenue']) - Decimal(before['paid_revenue']) == Decimal('960.54')
    assert after['paid_by_month']['2026-01'] == '960.54'
    assert after['paid_by_category']['Repairs'] == '325.50'
    assert Decimal(after['pending_amount']) - Decimal(before['pending_amount']) == Decimal('250.00')


def test_conflicting_duplicate_ids_exclude_every_occurrence():
    data = [['X','2026-01-01','Parts','90.00','Paid'],['X','2026-02-01','Parts','95.00','Paid']]
    valid, audit, report = analyze(records(data))
    assert valid == []
    assert report['paid_revenue'] == '0.00'
    assert report['duplicate_rows_excluded'] == 2
    assert all('Duplicate' in r['reason'] for r in audit)


@pytest.mark.parametrize('field,value,reason', [
    ('date','2026-02-30','Invalid ISO date'), ('date','01/02/2026','Invalid ISO date'),
    ('amount','10.001','Invalid amount'), ('amount','NaN','Invalid amount'),
    ('amount',True,'Invalid amount'), ('amount','1000000000','Invalid amount'),
    ('status','paid','Invalid status'), ('category','','Missing category'), ('sale_id','','Missing sale ID')])
def test_invalid_values_are_excluded_with_reason(field,value,reason):
    row = records([['S1','2026-02-28','Parts','10.01','Paid']])[0]
    row[field] = value
    valid,audit,report = analyze([row])
    assert not valid and report['rejected_records'] == 1
    assert audit[0]['reason'].startswith(reason)


def test_decimal_arithmetic_and_zero_and_negative():
    data = [[str(n),'2026-01-01','Parts',v,'Paid'] for n,v in enumerate(['0.10','0.20','-0.30','0'])]
    report = analyze(records(data))[2]
    assert report['paid_revenue'] == '0.00' and report['paid_count'] == 4
    assert report['negative_paid_count'] == 1


def test_headers_move_columns_reorder_and_blank_rows_skip(tmp_path):
    path = tmp_path / 'input.xlsx'
    save(path, [ADDED[0], [None]*5, ADDED[1]], list(reversed(HEADERS)), offset=3)
    data = read_sales(path)
    assert [r['source_row'] for r in data] == [5,7]
    assert analyze(data)[2]['paid_revenue'] == '140.25'


def test_repeated_or_missing_headers_fail(tmp_path):
    path = tmp_path / 'bad.xlsx'
    b = openpyxl.Workbook(); b.active.title = 'Sales'
    b.active.append(HEADERS+['amount']); b.save(path)
    with pytest.raises(ValueError, match='Repeated'): read_sales(path)
    b.active.delete_cols(1); b.save(path)
    with pytest.raises(ValueError, match='headers'): read_sales(path)


def test_source_formula_not_silently_trusted(tmp_path):
    path = tmp_path / 'input.xlsx'
    save(path,[['S1','2026-01-01','Parts','=10+5','Paid']])
    assert analyze(read_sales(path))[1][0]['reason'].startswith('Formula')


def test_rerun_replaces_rows_and_keeps_sources_unchanged(tmp_path):
    src, out = tmp_path / 'input space', tmp_path / 'output space'
    src.mkdir(); save(src/'input.xlsx', rows()+ADDED)
    original = (src/'input.xlsx').read_bytes()
    one = build(src,out); two = build(src,out)
    assert one == two and original == (src/'input.xlsx').read_bytes()
    book = openpyxl.load_workbook(out/'dashboard.xlsx', data_only=False)
    assert book['ValidSales'].max_row == 1301
    assert '$1301' in book['Dashboard']['A3'].value
    assert len(book['Dashboard']._charts) == 2
    assert [len(c.series[0].val.numRef.numCache.pt) for c in book['Dashboard']._charts] == [13, 7]


def test_empty_source_and_no_paid_rows(tmp_path):
    src, out = tmp_path/'input',tmp_path/'output'; src.mkdir()
    save(src/'input.xlsx',[])
    report=build(src,out)
    assert report['input_records'] == 0 and report['paid_revenue'] == '0.00'
    save(src/'input.xlsx',[['S1','2026-01-01','Parts','99','Pending']])
    report=build(src,out)
    assert report['paid_count'] == 0 and report['pending_amount'] == '99.00'


def test_output_must_not_overlap_input(tmp_path):
    with pytest.raises(ValueError,match='separate'): build(tmp_path,tmp_path/'output')


def test_literal_category_names_do_not_merge_case_or_wildcards(tmp_path):
    src, out = tmp_path/'input',tmp_path/'output'; src.mkdir()
    data = [[str(n),'2026-01-01',cat,amount,'Paid'] for n,(cat,amount) in enumerate([('Parts','10'),('parts','20'),('*','30'),('?','40')])]
    save(src/'input.xlsx',data)
    report=build(src,out)
    assert report['paid_by_category'] == {'*':'30.00','?':'40.00','Parts':'10.00','parts':'20.00'}
    # Exercise actual formula results when the container's recalculator is available.
    import shutil, subprocess
    if not shutil.which('xlrecalc'):
        pytest.skip('LibreOffice recalculator required')
    subprocess.run(['xlrecalc',str(out/'dashboard.xlsx'),str(out/'checked.xlsx')],check=True,capture_output=True)
    book=openpyxl.load_workbook(out/'checked.xlsx',data_only=True)
    assert [book['Dashboard'].cell(n,2).value for n in range(15,19)] == [30,40,10,20]


def test_error_cells_rejected(tmp_path):
    path=tmp_path/'input.xlsx'
    save(path,[['S1','2026-01-01','#VALUE!','10','Paid']])
    assert analyze(read_sales(path))[1][0]['reason'] == 'Invalid cell value'


def test_paid_and_pending_same_amount_stay_separate_and_cancelled_is_visible():
    data = [['P','2026-01-01','Parts','50.50','Paid'],
            ['U','2026-01-01','Parts','50.50','Pending'],
            ['C','2026-01-01','Parts','50.50','Cancelled']]
    valid,audit,report=analyze(records(data))
    assert len(valid) == len(audit) == 3
    assert report['paid_revenue'] == report['pending_amount'] == report['cancelled_amount_excluded'] == '50.50'
    assert report['paid_count'] == report['pending_count'] == report['cancelled_count'] == 1


def test_corrected_duplicate_enters_on_next_refresh(tmp_path):
    src,out=tmp_path/'input',tmp_path/'output'; src.mkdir()
    data=[['X','2026-01-01','Parts','70.00','Paid'],['X','2026-01-02','Parts','80.00','Paid']]
    save(src/'input.xlsx',data)
    assert build(src,out)['paid_revenue'] == '0.00'
    data[1][0]='Y'
    save(src/'input.xlsx',data)
    report=build(src,out)
    assert report['paid_revenue'] == '150.00' and report['rejected_records'] == 0
    assert report['paid_count'] == 2


def test_audit_text_cannot_become_an_active_spreadsheet_formula(tmp_path):
    import csv
    src,out=tmp_path/'input',tmp_path/'output'; src.mkdir()
    save(src/'input.xlsx',[['=1+2','2026-01-01','Parts','10','Paid']])
    build(src,out)
    with (out/'audit.csv').open() as f:
        row=next(csv.DictReader(f))
    assert row['sale_id'] == "'=1+2" and row['disposition'] == 'Rejected'
    book=openpyxl.load_workbook(out/'dashboard.xlsx',data_only=False)
    assert book['Audit']['B2'].data_type == 's'
