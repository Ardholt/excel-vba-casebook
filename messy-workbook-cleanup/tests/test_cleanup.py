import csv
import json
import shutil
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from normalize_records import FIELDS, parse_amount, parse_date, run, table
from generate_input import generate


def fixture_input(tmp_path, rows, order='DMY', headers=FIELDS, preamble=False):
    source = tmp_path / 'source'
    source.mkdir()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = table(wb, 'Records', headers, rows)
    if preamble:
        ws.insert_rows(1, 2)
        ws['A1'] = 'Customer register'
    wb.save(source / 'input.xlsx')
    (source / 'rules.json').write_text(json.dumps({'date_order': order}))
    return source


@pytest.mark.parametrize('order,expected', [('DMY', '2026-04-03'), ('MDY', '2026-03-04')])
def test_ambiguous_date_uses_declared_order(order, expected):
    assert parse_date('03/04/2026', order).isoformat() == expected


@pytest.mark.parametrize('raw,expected', [('£1,240.50', '1240.50'), ('(£85.20)', '-85.20'),
                                          ('($2,002.00)', '-2002.00'), (' €19.2 ', '19.20'), ('-0.01', '-0.01'), ('£-85.20', '-85.20'), ('1,200.00 €', '1200.00')])
def test_amounts(raw, expected):
    assert parse_amount(raw) == Decimal(expected)


@pytest.mark.parametrize('raw', ['12,34.50', '1.234,50', 'NaN', 'Infinity', '1.234', '(12.00', '$2,002.00)', '(-12)', '£$12', '', True])
def test_bad_amounts_rejected(raw):
    with pytest.raises(ValueError):
        parse_amount(raw)


@pytest.mark.parametrize('raw', ['31/02/2026', '2026-13-01', '04/03/26', 46000, datetime(2026, 1, 1, 12)])
def test_bad_dates_rejected(raw):
    with pytest.raises(ValueError):
        parse_date(raw, 'DMY')


def test_business_result_and_money(tmp_path):
    generate(tmp_path / 'source')
    out = tmp_path / 'result'
    report = run(tmp_path / 'source', out)
    assert (report['input_records'], report['kept_records'], report['duplicate_records'], report['rejected_records']) == (550, 472, 65, 13)
    assert report['kept_amount'] == '194075.70'
    assert report['duplicate_amount'] == '26183.98'
    assert report['valid_source_amount'] == '220259.68'
    with (out / 'audit.csv').open() as f:
        audit = list(csv.DictReader(f))
    assert [int(r['source_row']) for r in audit] == list(range(2, 552))
    assert {r['row_id'] for r in audit if r['status'] == 'rejected'} == {'R004', 'R028', 'R029'} | {f'R{i:03}' for i in range(541, 551)}
    with (out / 'duplicates.csv').open() as f:
        mapping = {r['row_id']: r['kept_row_id'] for r in csv.DictReader(f)}
    assert mapping == {'R002': 'R001', 'R024': 'R005', 'R025': 'R008', 'R026': 'R012', 'R027': 'R017',
                       **{f'R{481+i:03}': f'R{31+i*7:03}' for i in range(60)}}
    rows = list(openpyxl.load_workbook(out / 'customer-register.xlsx')['Records'].values)
    assert rows[1][1] == '00071'
    assert rows[1][3] == datetime(2026, 4, 3)
    assert rows[2][4] == -85.2
    assert next(row for row in rows[1:] if row[0] == 'R030')[1] == '00330'  # Same name/date/amount, different identity remains.


def test_repeat_and_cleaned_input_are_stable(tmp_path):
    generate(tmp_path / 'source')
    a = run(tmp_path / 'source', tmp_path / 'result')
    b = run(tmp_path / 'source', tmp_path / 'result')
    assert a['business_result_sha256'] == b['business_result_sha256']
    next_source = tmp_path / 'second source'
    next_source.mkdir()
    shutil.copyfile(tmp_path / 'result' / 'customer-register.xlsx', next_source / 'input.xlsx')
    shutil.copyfile(tmp_path / 'source' / 'rules.json', next_source / 'rules.json')
    c = run(next_source, tmp_path / 'next result')
    assert c['business_result_sha256'] == a['business_result_sha256']
    assert c['duplicate_records'] == c['rejected_records'] == c['field_changes'] == 0
    assert c['kept_records'] == a['kept_records']


def test_reordered_headers_below_title_and_blank_rows(tmp_path):
    headers = ('amount', 'name', 'customer_id', 'date', 'row_id')
    source = fixture_input(tmp_path, [['£12.30', '  Unit   A ', '001', '2026-02-28', 'A'],
                                      [None] * 5, ['12.30', 'unit a', '001', '28/02/2026', 'B']], headers=headers, preamble=True)
    result = run(source, tmp_path / 'result')
    assert (result['input_records'], result['kept_records'], result['duplicate_records']) == (2, 1, 1)
    rows = list(openpyxl.load_workbook(tmp_path / 'result' / 'customer-register.xlsx')['Records'].values)
    assert rows[1][1:3] == ('001', 'Unit A')
    assert rows[1][-1] == 4


def test_rejection_preserves_raw_and_all_reasons(tmp_path):
    source = fixture_input(tmp_path, [['R1', '0001', 'A', '31/02/2026', 'oops'],
                                      ['R2', 1, 'A', '2026-02-01', 2],
                                      ['R3', '0003', 'A', '2026-02-01', '=1+1']])
    result = run(source, tmp_path / 'result')
    assert result['rejected_records'] == 3
    wb = openpyxl.load_workbook(tmp_path / 'result' / 'customer-register.xlsx')
    rows = list(wb['Rejected'].values)
    assert rows[1][4:6] == ('31/02/2026', 'oops')
    assert rows[1][-1] == 'invalid date; invalid amount'
    assert rows[3][-2] == '=1+1'
    assert wb['Rejected']['F4'].data_type == 's'


def test_no_fuzzy_merge_or_id_case_merge(tmp_path):
    source = fixture_input(tmp_path, [['A', 'ab01', 'Unit N-01', '2026-01-01', 10],
                                      ['B', 'AB01', 'Unit N-01', '2026-01-01', 10],
                                      ['C', 'ab01', 'Unit N-0', '2026-01-01', 10],
                                      ['D', 'ab01', ' Unit  n-01 ', '01/01/2026', '10.00']])
    result = run(source, tmp_path / 'result')
    assert result['kept_records'] == 3
    assert result['duplicate_records'] == 1


def test_missing_schema_and_invalid_rules_fail(tmp_path):
    source = fixture_input(tmp_path, [['A']], headers=['row_id'])
    with pytest.raises(ValueError, match='headers'):
        run(source, tmp_path / 'out')
    (source / 'rules.json').write_text('{"date_order":"guess"}')
    with pytest.raises(ValueError, match='date_order'):
        run(source, tmp_path / 'out')


def test_source_unchanged_and_nested_output_refused(tmp_path):
    source = fixture_input(tmp_path, [['A', '0001', 'A', '2026-01-01', '0.00']])
    before = (source / 'input.xlsx').read_bytes()
    run(source, tmp_path / 'out')
    assert before == (source / 'input.xlsx').read_bytes()
    with pytest.raises(ValueError, match='separate'):
        run(source, source / 'out')


def test_empty_and_duplicate_header(tmp_path):
    source = fixture_input(tmp_path, [])
    result = run(source, tmp_path / 'out')
    assert result['input_records'] == result['kept_records'] == 0
    wb = openpyxl.load_workbook(source / 'input.xlsx')
    wb['Records']['F1'] = 'amount'
    wb.save(source / 'input.xlsx')
    with pytest.raises(ValueError, match='duplicate required header'):
        run(source, tmp_path / 'out')


def test_replacement_input_changes_result_and_row_ids_need_not_be_unique(tmp_path):
    source = fixture_input(tmp_path, [['same', '0001', 'Unit A', '03/04/2026', '10.00'],
                                      ['same', '0002', 'Unit B', '03/04/2026', '-2.50']], order='MDY')
    out = tmp_path / 'result'
    report = run(source, out)
    assert report['kept_records'] == 2
    assert report['kept_amount'] == '7.50'
    rows = list(openpyxl.load_workbook(out / 'customer-register.xlsx')['Records'].values)
    assert rows[1][3] == rows[2][3] == datetime(2026, 3, 4)
    wb = openpyxl.load_workbook(source / 'input.xlsx')
    wb['Records'].append(['new', '0003', 'Unit C', '03/05/2026', '11.11'])
    wb.save(source / 'input.xlsx')
    changed = run(source, out)
    assert changed['kept_records'] == 3
    assert changed['kept_amount'] == '18.61'
    assert changed['business_result_sha256'] != report['business_result_sha256']


def test_identifiers_only_trim_and_missing_sheet_fails(tmp_path):
    source = fixture_input(tmp_path, [['A', '00  01', 'A', '2026-01-01', 1],
                                      ['B', '00 01', 'A', '2026-01-01', 1]])
    assert run(source, tmp_path / 'out')['kept_records'] == 2
    wb = openpyxl.load_workbook(source / 'input.xlsx')
    wb['Records'].title = 'Other'
    wb.save(source / 'input.xlsx')
    with pytest.raises(ValueError, match='Records sheet'):
        run(source, tmp_path / 'out')


def test_leap_date_and_records_with_extra_cells(tmp_path):
    assert parse_date('29/02/2024', 'DMY').isoformat() == '2024-02-29'
    with pytest.raises(ValueError):
        parse_date('29/02/2026', 'DMY')
    source = fixture_input(tmp_path, [['A', '0001', 'A', '2026-01-01', 1, 'note'],
                                      [None, None, None, None, None, 'orphan note']], headers=(*FIELDS, 'notes'))
    result = run(source, tmp_path / 'out')
    assert result['input_records'] == 2
    assert result['kept_records'] == result['rejected_records'] == 1


def test_money_reconciles_exactly_to_cents(tmp_path):
    source = fixture_input(tmp_path, [['A', '01', 'A', '2026-01-01', '0.10'],
                                      ['B', '02', 'B', '2026-01-01', '0.20'],
                                      ['C', '03', 'C', '2026-01-01', '(0.30)'],
                                      ['D', '02', 'b', '01/01/2026', '0.20']])
    report = run(source, tmp_path / 'out')
    assert report['kept_amount'] == '0.00'
    assert report['duplicate_amount'] == report['valid_source_amount'] == '0.20'
    assert all(report['checks'].values())
