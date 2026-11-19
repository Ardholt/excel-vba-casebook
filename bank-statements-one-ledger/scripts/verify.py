"""Independent output conservation and workbook serialization checks."""
import csv
import json
from decimal import Decimal
from pathlib import Path
from openpyxl import load_workbook


def verify(output, checked):
    output=Path(output)
    report=json.loads((output/'report.json').read_text())
    with (output/'ledger.csv').open(newline='') as f: ledger=list(csv.DictReader(f))
    with (output/'reconciliation.csv').open(newline='') as f: accounts=list(csv.DictReader(f))
    assert len(ledger)==report['output_records']
    assert report['input_records']==len(ledger)+report['excluded_records']
    assert all(report['checks'].values())
    assert len({(r['source_file'],r['source_row']) for r in ledger})==len(ledger)
    total=Decimal('0')
    for a in accounts:
        rows=[r for r in ledger if r['account']==a['account']]
        assert len(rows)==int(a['ledger_records'])
        movement=sum((Decimal(r['amount']) for r in rows),Decimal('0'))
        total+=movement
        assert movement==Decimal(a['credits'])-Decimal(a['debits'])
        if a['status']!='INCOMPLETE':
            balance=Decimal(a['opening_balance'])
            for r in rows:
                balance+=Decimal(r['amount'])
                assert balance==Decimal(r['calculated_balance'])
                difference=Decimal(r['reported_balance'])-balance
                assert difference==Decimal(r['difference'])
                assert r['check']==('OK' if difference==0 else 'MISMATCH')
            assert balance==Decimal(a['calculated_close'])
            assert Decimal(a['declared_close'])-balance==Decimal(a['closing_difference'])
    assert total==Decimal(report['net_movement'])
    before=load_workbook(output/'ledger.xlsx',data_only=True)
    after=load_workbook(checked,data_only=True)
    cell_count=0
    assert before.sheetnames==after.sheetnames
    for name in before.sheetnames:
        for row in before[name]:
            for cell in row:
                expected=cell.value; actual=after[name][cell.coordinate].value
                if expected in ('',None): assert actual in ('',None)
                elif isinstance(expected,(int,float)):
                    assert Decimal(str(expected))==Decimal(str(actual))
                else: assert expected==actual
                if expected not in ('',None): cell_count+=1
    return {'records_and_money_conserved':True,'running_and_closing_checks_verified':True,
            'libreoffice_cell_values_match':True,'nonempty_cells_compared':cell_count}
