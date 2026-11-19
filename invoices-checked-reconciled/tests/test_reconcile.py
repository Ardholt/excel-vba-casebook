import hashlib
import json
import os
from decimal import Decimal
from pathlib import Path
import sys
import pytest
from openpyxl import Workbook, load_workbook
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from reconcile import reconcile, write_outputs, HEADERS, decimal
from run_pipeline import verify_workbook

ROOT = Path(__file__).resolve().parents[1]


def source(tmp_path, lines=None, invoices=None, payments=None, reordered=False):
    book = Workbook()
    book.remove(book.active)
    records = {'Lines': lines if lines is not None else [['I-1','L-1','1','10.00','0.20']],
               'Invoices': invoices if invoices is not None else [['I-1','12.00']],
               'Payments': payments if payments is not None else []}
    for name, headers in HEADERS.items():
        ws = book.create_sheet(name)
        if reordered:
            ws.append(['Batch for reconciliation'])
            ws.append([])
            ws.append(list(reversed(headers)) + ['extra'])
            for row in records[name]:
                ws.append(list(reversed(row)) + ['ignored'])
        else:
            ws.append(headers)
            for row in records[name]:
                ws.append(row)
    path = tmp_path / 'input.xlsx'
    book.save(path)
    return path


@pytest.mark.parametrize('quantity,price,rate,net,tax,gross', [
    ('1','1.005','0.20','1.01','0.20','1.21'),
    ('-1','1.005','0.20','-1.01','-0.20','-1.21'),
    ('1','0.10','0.05','0.10','0.01','0.11'),
    ('-1','0.10','0.05','-0.10','-0.01','-0.11'),
    ('2.5','3.33','0','8.33','0.00','8.33'),
])
def test_half_up_line_rounding(tmp_path, quantity, price, rate, net, tax, gross):
    r = reconcile(source(tmp_path, lines=[['I-1','L-1',quantity,price,rate]], invoices=[['I-1',gross]]))
    assert [r['invoices'][0][k] for k in ('net','tax','gross')] == list(map(Decimal, (net,tax,gross)))


def test_round_each_line_before_sum(tmp_path):
    result = reconcile(source(tmp_path, lines=[['I-1',f'L-{n}','1','0.025','0.20'] for n in range(3)], invoices=[['I-1','0.12']]))
    assert result['invoices'][0]['net'] == Decimal('0.09')
    assert result['invoices'][0]['tax'] == Decimal('0.03')
    assert result['invoices'][0]['gross'] == Decimal('0.12')


@pytest.mark.parametrize('amount,balance,status', [('5.00','7.00','DUE'),('12.00','0.00','SETTLED'),('15.00','-3.00','CREDIT'),('-2.00','14.00','DUE')])
def test_signed_payment_balances(tmp_path, amount, balance, status):
    result = reconcile(source(tmp_path, payments=[['P1','I-1',amount]]))
    assert result['invoices'][0]['balance'] == Decimal(balance)
    assert result['invoices'][0]['balance_status'] == status


def test_credit_note_refund(tmp_path):
    result = reconcile(source(tmp_path, lines=[['I-1','L1','-2','10','0.20']], invoices=[['I-1','-24']], payments=[['R1','I-1','-10']]))
    assert result['invoices'][0]['balance'] == Decimal('-14')


@pytest.mark.parametrize('second', ['12.00','22.50'])
def test_all_duplicate_id_occurrences_held(tmp_path, second):
    result = reconcile(source(tmp_path, payments=[['P1','I-1','12'],['P1','I-1',second],['P2','I-1','2']]))
    assert result['invoices'][0]['allocated'] == Decimal('2')
    assert result['report']['payment_rows']['duplicate_id'] == 2
    assert result['report']['payment_amounts']['duplicate_id'] == str(Decimal('12.00') + Decimal(second))
    assert all(result['report']['checks'].values())


def test_unmatched_payment_is_not_lost(tmp_path):
    result = reconcile(source(tmp_path, payments=[['P1','unknown','7.50']]))
    assert result['report']['payment_amounts']['unallocated'] == '7.50'
    assert result['invoices'][0]['allocated'] == 0
    assert result['exceptions'][0]['reason'] == 'unallocated_payment'


def test_distinct_equal_payments_both_allocate(tmp_path):
    result = reconcile(source(tmp_path, payments=[['P1','I-1','6'],['P2','I-1','6']]))
    assert result['invoices'][0]['balance'] == 0


def test_declared_discrepancy_visible_not_plugged(tmp_path):
    result = reconcile(source(tmp_path, invoices=[['I-1','13.25']]))
    assert result['invoices'][0]['gross'] == Decimal('12')
    assert result['invoices'][0]['difference'] == Decimal('1.25')
    assert result['report']['declared_mismatches'] == 1


def test_header_positions_and_order_do_not_matter(tmp_path):
    result = reconcile(source(tmp_path, reordered=True, payments=[['00007','I-1','12']]))
    assert result['invoices'][0]['balance'] == 0
    assert result['payments'][0]['payment_id'] == '00007'
    assert result['payments'][0]['source_row'] == 4


@pytest.mark.parametrize('value', ['bad', 'NaN', 'Infinity', '-0.01', '1.01'])
def test_invalid_tax_holds_entire_invoice(tmp_path, value):
    result = reconcile(source(tmp_path, lines=[['I-1','L1','1','10',value],['I-1','L2','1','10','0.2']], payments=[['P1','I-1','12']]))
    assert result['invoices'] == []
    assert result['payments'][0]['status'] == 'unallocated'
    assert result['report']['held_invoice_rows'] == 1
    assert all(result['report']['checks'].values())


def test_duplicate_line_holds_invoice(tmp_path):
    result = reconcile(source(tmp_path, lines=[['I-1','L1','1','10','0.2']] * 2))
    assert result['invoices'] == []
    assert result['report']['held_line_rows'] == 2


def test_duplicate_invoice_headers_held(tmp_path):
    result = reconcile(source(tmp_path, invoices=[['I-1','12'], ['I-1','12']]))
    assert result['invoices'] == []
    assert result['report']['held_invoice_rows'] == 2
    assert all(result['report']['checks'].values())


def test_invalid_payment_preserved_and_not_allocated(tmp_path):
    result = reconcile(source(tmp_path, payments=[['P1','I-1','oops'],['P2','I-1','1.001']]))
    assert result['report']['payment_rows']['invalid'] == 2
    assert result['report']['unparseable_payment_rows'] == 2
    assert result['invoices'][0]['allocated'] == 0


def test_input_formulas_rejected(tmp_path):
    result = reconcile(source(tmp_path, lines=[['I-1','L1','=1+1','10','0.2']]))
    assert not result['invoices']
    assert result['exceptions'][0]['reason'] == 'invalid_line'


def test_no_lines_holds_invoice(tmp_path):
    result = reconcile(source(tmp_path, lines=[]))
    assert result['report']['held_invoice_rows'] == 1


def test_missing_headers_fail_clearly(tmp_path):
    path = source(tmp_path)
    book = load_workbook(path)
    book['Lines']['C1'] = 'wrong'
    book.save(path)
    with pytest.raises(ValueError, match='Missing required headers'):
        reconcile(path)


def test_repeat_run_same_business_result_and_source_bytes(tmp_path):
    path = source(tmp_path, payments=[['P1','I-1','12']])
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    one = reconcile(path)
    write_outputs(one, tmp_path / 'result')
    before = (tmp_path / 'result/invoices.csv').read_bytes()
    two = reconcile(path)
    write_outputs(two, tmp_path / 'result')
    assert one == two
    assert (tmp_path / 'result/invoices.csv').read_bytes() == before
    assert hashlib.sha256(path.read_bytes()).hexdigest() == digest


def test_batch_independent_integer_cents_check():
    # Independent oracle: rational arithmetic, with integer half-away rounding.
    from fractions import Fraction
    def cents(value):
        value *= 100
        sign = -1 if value < 0 else 1
        value = abs(value)
        return sign * ((2 * value.numerator + value.denominator) // (2 * value.denominator))
    path = ROOT / 'input/input.xlsx'
    book = load_workbook(path, data_only=True)
    totals = {}
    for iid, lid, quantity, price, rate in book['Lines'].iter_rows(min_row=2, values_only=True):
        net = cents(Fraction(str(quantity)) * Fraction(str(price)))
        tax = cents(Fraction(net, 100) * Fraction(str(rate)))
        totals[iid] = totals.get(iid, 0) + net + tax
    result = reconcile(path)
    assert len(totals) == 120
    assert result['report']['input_records']['Lines'] == 360
    assert {r['invoice_id']: int(r['gross'] * 100) for r in result['invoices']} == totals
    assert result['report']['declared_mismatches'] == 7
    assert result['report']['payment_rows']['duplicate_id'] == 6
    assert result['report']['payment_rows']['unallocated'] == 4
    book.close()


def test_recalculated_workbook_matches_current_input():
    src = Path(os.environ.get('CASE_INPUT', ROOT / 'input')) / 'input.xlsx'
    out = Path(os.environ.get('CASE_RESULT', ROOT / 'output'))
    result = reconcile(src)
    assert verify_workbook(out / 'checked-ledger.xlsx', result) > 0 or not result['invoices']
    report = json.loads((out / 'report.json').read_text())
    assert report['totals'] == result['report']['totals']
    assert all(report['checks'].values())
