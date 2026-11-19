import csv
import json
import shutil
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from ledger import build, compute, pennies, read_date

BASE = Path(__file__).resolve().parents[1]

@pytest.fixture
def inputs(tmp_path):
    target = tmp_path / 'input with spaces'
    shutil.copytree(BASE / 'input', target)
    return target


def rewrite(path, transform):
    with path.open(newline='') as f: rows = list(csv.reader(f))
    rows = transform(rows)
    with path.open('w', newline='') as f: csv.writer(f).writerows(rows)


def test_all_records_and_money_conserved(inputs):
    ledger, recon, rejects, issues, report = compute(inputs)
    assert len(ledger) == 810 and not rejects
    assert all(a['input_records'] == 270 for a in recon)
    assert all(a['period_start'] == '2026-10-01' and a['period_end'] == '2026-12-31' for a in recon)
    assert report['input_records'] == report['output_records'] == 810
    assert report['total_credits'] == '54544.50'
    assert report['total_debits'] == '26264.68'
    assert report['net_movement'] == '28279.82'
    assert [a['calculated_close'] for a in recon] == ['10808.28','9817.04','9904.50']
    assert len({(r['source_file'],r['source_row']) for r in ledger}) == 810
    assert sum(pennies(r['amount']) for r in ledger) == 2827982


def test_known_difference_is_not_plugged(inputs):
    ledger, recon, rejects, issues, report = compute(inputs)
    assert ledger[-1]['amount'] == '0.00'
    assert ledger[-1]['reported_balance'] == '9904.25'
    assert ledger[-1]['calculated_balance'] == '9904.50'
    assert ledger[-1]['difference'] == '-0.25'
    assert report['running_balance_mismatches'] == report['closing_balance_mismatches'] == 1
    assert [a['status'] for a in recon] == ['OK','OK','REVIEW']
    assert [(i['check'],i['difference']) for i in issues] == [('running_balance','-0.25'),('closing_balance','-0.25')]
    assert report['result'] == 'REVIEW_REQUIRED'


def test_equal_cross_account_payments_remain(inputs):
    ledger, *_ = compute(inputs)
    receipts = [r for r in ledger if r['amount']=='450.00']
    assert {r['account'] for r in receipts} == {'ACCT-A','ACCT-B'}
    assert {r['date'] for r in receipts} == {'2026-10-01'}


def test_same_day_order_and_charges(inputs):
    ledger, *_ = compute(inputs)
    assert [(r['id'],r['source_row']) for r in ledger[:2]] == [('A001',2),('A002',3)]
    assert ledger[0]['date'] == ledger[1]['date']
    assert ledger[2]['amount'] == '-10.00'
    assert ledger[5]['amount'] == ledger[275]['amount'] == '-0.10'


def test_headers_move_and_reorder(inputs):
    rewrite(inputs/'layout-a.csv', lambda rows: [['Statement A'],[]] + [list(reversed(row)) for row in rows])
    book=load_workbook(inputs/'layout-b.xlsx'); ws=book['Transactions']; ws.insert_rows(1,3)
    book.save(inputs/'layout-b.xlsx')
    ledger,*_=compute(inputs)
    assert ledger[0]['source_row']==4 and ledger[270]['source_row']==5
    assert ledger[0]['amount']=='450.00' and ledger[270]['amount']=='450.00'


def test_repeat_run_and_sources_unchanged(inputs,tmp_path):
    before={p.name:p.read_bytes() for p in inputs.iterdir()}
    first=build(inputs,tmp_path/'output with spaces')
    second=build(inputs,tmp_path/'output with spaces')
    first.pop('runtime_seconds'); second.pop('runtime_seconds')
    assert first==second
    assert before=={p.name:p.read_bytes() for p in inputs.iterdir()}
    book=load_workbook(tmp_path/'output with spaces'/'ledger.xlsx',data_only=True)
    assert book['Ledger'].max_row==811
    assert book['Ledger']['G811'].value==0
    assert book['Ledger']['H811'].value==9904.25
    assert book['Ledger']['I811'].value==9904.50


def test_invalid_row_makes_following_balances_unverified(inputs):
    def change(rows): rows[2][3]='oops'; return rows
    rewrite(inputs/'layout-a.csv',change)
    ledger,recon,rejected,_,report=compute(inputs)
    assert len(ledger)==809 and len(rejected)==1
    assert report['checks']['records_conserved']
    assert recon[0]['status']=='INCOMPLETE'
    assert all(r['check']=='UNVERIFIED' and r['calculated_balance']=='' for r in [r for r in ledger[1:] if r['account']=='ACCT-A'])
    assert report['result']=='INPUT_ERRORS'


def test_duplicate_identifier_rejected_not_double_counted(inputs):
    rewrite(inputs/'layout-a.csv',lambda rows: rows+[rows[1]])
    ledger,recon,rejected,_,report=compute(inputs)
    assert len(ledger)==810 and report['input_records']==811
    assert rejected[0]['reason']=='duplicate transaction identifier within account'
    assert recon[0]['status']=='INCOMPLETE'


def test_blank_debit_credit_are_zero(inputs):
    def change(rows): rows[1][3]=''; rows[1][4]=''; return rows
    rewrite(inputs/'layout-a.csv',change)
    ledger,*_=compute(inputs)
    assert ledger[0]['amount']=='0.00'
    assert ledger[0]['check']=='MISMATCH'

@pytest.mark.parametrize('bad', ['1.001','NaN','Infinity','£1.00','1,000.00','',True])
def test_invalid_money_is_never_rounded_or_guessed(bad):
    with pytest.raises(ValueError): pennies(bad)

@pytest.mark.parametrize('direction,amount', [('SIDE','1.00'),('IN','-1.00')])
def test_unsigned_direction_layout_validation(inputs,direction,amount):
    def change(rows): rows[1][3:5]=[amount,direction]; return rows
    rewrite(inputs/'layout-c.csv',change)
    _,recon,rejected,_,_=compute(inputs)
    assert rejected[0]['source_file']=='layout-c.csv'
    assert recon[2]['status']=='INCOMPLETE'


def test_intermediate_discrepancy_even_when_close_matches(inputs):
    def change(rows): rows[2][5]='1324.49'; return rows
    rewrite(inputs/'layout-a.csv',change)
    _,recon,_,issues,_=compute(inputs)
    assert recon[0]['closing_difference']=='0.00'
    assert recon[0]['running_mismatches']==1
    assert recon[0]['status']=='REVIEW'
    assert issues[0]['source_row']==3 and issues[0]['difference']=='-0.01'


def test_empty_account_uses_opening_and_closing(inputs):
    rewrite(inputs/'layout-a.csv',lambda rows: rows[:1])
    settings=json.loads((inputs/'accounts.json').read_text()); settings['layout-a.csv']['closing_balance']='1000.00'
    (inputs/'accounts.json').write_text(json.dumps(settings))
    _,recon,_,_,report=compute(inputs)
    assert recon[0]['status']=='OK' and recon[0]['ledger_records']==0
    assert report['input_records']==540


def test_date_order_is_explicit():
    assert read_date('03/04/2026','DMY')=='2026-04-03'
    assert read_date('03/04/2026','MDY')=='2026-03-04'
    with pytest.raises(ValueError): read_date('31/02/2026','DMY')


def test_missing_schema_fails(inputs):
    rewrite(inputs/'layout-a.csv', lambda rows: [['wrong']]+rows[1:])
    with pytest.raises(ValueError,match='headers'): compute(inputs)


def test_account_codes_cannot_merge(inputs):
    settings=json.loads((inputs/'accounts.json').read_text()); settings['layout-b.xlsx']['account']='ACCT-A'
    (inputs/'accounts.json').write_text(json.dumps(settings))
    with pytest.raises(ValueError,match='distinct'): compute(inputs)


def test_input_output_overlap_rejected(inputs):
    with pytest.raises(ValueError,match='separate'): build(inputs,inputs/'output')


def test_description_formula_is_inert(inputs,tmp_path):
    def change(rows): rows[1][2]='=1+1'; return rows
    rewrite(inputs/'layout-a.csv',change)
    build(inputs,tmp_path/'out')
    book=load_workbook(tmp_path/'out'/'ledger.xlsx')
    assert book['Ledger']['F2'].value=='=1+1' and book['Ledger']['F2'].data_type=='s'


def test_correcting_declared_and_reported_balance_clears_review(inputs):
    def change(rows): rows[-1][-1]='9904.50'; return rows
    rewrite(inputs/'layout-c.csv',change)
    settings=json.loads((inputs/'accounts.json').read_text()); settings['layout-c.csv']['closing_balance']='9904.50'
    (inputs/'accounts.json').write_text(json.dumps(settings))
    _,recon,rejected,issues,report=compute(inputs)
    assert not rejected and not issues
    assert report['result']=='BALANCED'
    assert all(a['status']=='OK' for a in recon)


def test_declared_close_checked_even_if_final_running_balance_passes(inputs):
    settings=json.loads((inputs/'accounts.json').read_text()); settings['layout-a.csv']['closing_balance']='10809.28'
    (inputs/'accounts.json').write_text(json.dumps(settings))
    _,recon,_,issues,_=compute(inputs)
    assert recon[0]['running_mismatches']==0
    assert recon[0]['closing_difference']=='1.00'
    assert recon[0]['status']=='REVIEW'
    assert any(i['account']=='ACCT-A' and i['check']=='closing_balance' for i in issues)


def test_statement_order_not_date_sort(inputs):
    def change(rows): rows[1][1]='05/10/2026'; return rows
    rewrite(inputs/'layout-a.csv',change)
    ledger,recon,*_=compute(inputs)
    assert ledger[0]['date']=='2026-10-05'
    assert ledger[1]['date']=='2026-10-01'
    assert recon[0]['status']=='OK'


def test_all_empty_statements_build_without_inventing_rows(inputs,tmp_path):
    rewrite(inputs/'layout-a.csv',lambda rows: rows[:1])
    rewrite(inputs/'layout-c.csv',lambda rows: rows[:1])
    book=load_workbook(inputs/'layout-b.xlsx'); ws=book['Transactions']; ws.delete_rows(2,ws.max_row)
    book.save(inputs/'layout-b.xlsx')
    settings=json.loads((inputs/'accounts.json').read_text())
    for a in settings.values(): a['closing_balance']=a['opening_balance']
    (inputs/'accounts.json').write_text(json.dumps(settings))
    report=build(inputs,tmp_path/'out')
    assert report['input_records']==report['output_records']==0
    assert report['result']=='BALANCED'


def test_source_money_independently_matches_full_quarter(inputs):
    from decimal import Decimal
    source_amounts=[]
    with (inputs/'layout-a.csv').open(newline='') as f:
        source_amounts += [Decimal(r['credit'] or '0')-Decimal(r['debit'] or '0') for r in csv.DictReader(f)]
    book=load_workbook(inputs/'layout-b.xlsx',data_only=True)
    source_amounts += [Decimal(str(row[3])) for row in list(book['Transactions'].values)[1:]]
    with (inputs/'layout-c.csv').open(newline='') as f:
        source_amounts += [Decimal(r['amount'])*(1 if r['direction']=='IN' else -1) for r in csv.DictReader(f)]
    ledger,_,_,_,report=compute(inputs)
    assert len(source_amounts)==len(ledger)==810
    assert sum(source_amounts)==sum(Decimal(r['amount']) for r in ledger)==Decimal('28279.82')
    assert sum(x for x in source_amounts if x>0)==Decimal(report['total_credits'])
    assert -sum(x for x in source_amounts if x<0)==Decimal(report['total_debits'])
