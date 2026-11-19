import csv
import sys
from pathlib import Path
import pytest
import openpyxl
sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'scripts'))
from customers import HEADERS, workbook, run, propose, resolve, read_customers


def customer(key, name='Cedar account', area='AREA-A'):
    return dict(customer_id=key, name=name, postcode=area, source_row=2)


def create(tmp_path, rows, pairs=None):
    source = tmp_path/'incoming files'
    source.mkdir()
    workbook(source/'input.xlsx', [('Customers',HEADERS,rows)])
    if pairs is not None:
        with (source/'confirmed.csv').open('w',newline='') as f:
            w=csv.writer(f); w.writerow(['source_id','target_id']); w.writerows(pairs)
    return source


def test_bundled_approval_conserves_every_id(tmp_path):
    root = Path(__file__).resolve().parents[1]
    report=run(root/'input',tmp_path/'result')
    assert (report['input_records'], report['retained_records'], report['merged_records']) == (480,440,40)
    book=openpyxl.load_workbook(tmp_path/'result/result.xlsx')
    audit={r[0]:r[1] for r in list(book['Audit'].values)[1:]}
    assert len(audit)==480 and set(audit.values())=={r[0] for r in list(book['Customers'].values)[1:]}
    assert audit['C0041']=='C0041' and audit['C0441']=='C0441'
    assert report['conservation'] and report['confirmation_errors']==0


def test_no_confirmation_even_exact_names_stay_separate(tmp_path):
    source=create(tmp_path,[['0001',' Cedar  account ','AREA-A'],['0002','CEDAR ACCOUNT','area-a']])
    report=run(source,tmp_path/'result')
    assert report['proposed_pairs']==1
    assert report['retained_records']==2 and report['merged_records']==0


def test_compatible_chain_keeps_final_target_and_original_fields(tmp_path):
    source=create(tmp_path,[['A','Old label','X'],['B','Middle label','Y'],['C','Chosen label','Z']],[['A','B'],['B','C']])
    report=run(source,tmp_path/'result')
    book=openpyxl.load_workbook(tmp_path/'result/result.xlsx')
    assert list(book['Customers'].values)[1:]==[('C','Chosen label','Z')]
    audit=list(book['Audit'].values)[1:]
    assert audit[0][2]=='A > B > C' and audit[0][5]=='Old label'
    assert report['merged_records']==2


@pytest.mark.parametrize('pairs,reason',[
    ([('A','Z')],'Unknown'),
    ([('A','B'),('A','C')],'Conflicting'),
    ([('A','B'),('B','A')],'Cycle'),
    ([('A','A')],'Self'),
    ([('','B')],'Unknown'),
    ([('A','B'),('B','C'),('C','B')],'Cycle'),
])
def test_invalid_batch_never_partly_merges(tmp_path,pairs,reason):
    source=create(tmp_path,[[k,'Cedar account','X'] for k in 'ABCD'],pairs+[('D','C')])
    report=run(source,tmp_path/'result')
    assert report['status']=='REJECTED' and report['retained_records']==4 and report['merged_records']==0
    assert reason in (tmp_path/'result/confirmation-errors.csv').read_text()


def test_repeated_confirmation_is_idempotent():
    mapping,_,errors=resolve([customer('A'),customer('B')],[('A','B'),('A','B')])
    assert not errors and mapping=={'A':'B','B':'B'}


def test_rerun_uses_immutable_sources_and_does_not_accumulate(tmp_path):
    source=create(tmp_path,[['A','Cedar account','X'],['B','Cedar acount','X']],[['B','A']])
    before=(source/'input.xlsx').read_bytes()
    out=tmp_path/'result'
    first=run(source,out)
    first_rows=list(openpyxl.load_workbook(out/'result.xlsx')['Audit'].values)
    second=run(source,out)
    assert first_rows==list(openpyxl.load_workbook(out/'result.xlsx')['Audit'].values)
    assert first['retained_records']==second['retained_records']==1
    assert before==(source/'input.xlsx').read_bytes()


def test_headers_can_move_and_reorder(tmp_path):
    path=tmp_path/'input.xlsx'
    workbook(path,[('Customers',['postcode','name','customer_id'],[['X','Cedar','0001']])])
    book=openpyxl.load_workbook(path); book['Customers'].insert_rows(1,2); book['Customers']['A1']='Customer register'; book.save(path)
    assert read_customers(path)[0]['customer_id']=='0001'


@pytest.mark.parametrize('rows',[[['A','Cedar','X'],['A','Birch','Y']],[['','Cedar','X']],[['A','','X']]])
def test_invalid_customer_identity_rejected(tmp_path,rows):
    source=create(tmp_path,rows)
    with pytest.raises(ValueError): run(source,tmp_path/'result')


def test_numeric_id_rejected_to_protect_leading_zeros(tmp_path):
    source=create(tmp_path,[[12,'Cedar','X']])
    with pytest.raises(ValueError,match='text'): run(source,tmp_path/'result')


def test_blank_or_different_area_never_proposes():
    assert propose([customer('A',area=''),customer('B',area='')])==[]
    assert propose([customer('A',area='X'),customer('B',area='Y')])==[]


def test_close_spelling_proposes_with_reason():
    result=propose([customer('A','Cedar account'),customer('B','Cedar acount')])
    assert len(result)==1 and result[0][5]=='Same area; close name spelling'


def test_no_same_or_nested_output_folder(tmp_path):
    source=create(tmp_path,[['A','Cedar','X']])
    with pytest.raises(ValueError,match='separate'): run(source,source)
    with pytest.raises(ValueError,match='separate'): run(source,source/'output')


def test_formula_like_names_are_literal_text(tmp_path):
    source=create(tmp_path,[['A','=2+2','X']])
    run(source,tmp_path/'result')
    cell=openpyxl.load_workbook(tmp_path/'result/result.xlsx')['Customers']['B2']
    assert cell.data_type=='s' and cell.value=='=2+2'


def test_input_error_replaces_stale_pass_report(tmp_path):
    source=create(tmp_path,[['A','Cedar','X']])
    out=tmp_path/'result'
    run(source,out)
    workbook(source/'input.xlsx',[('Customers',HEADERS,[['A','Cedar','X'],['A','Other','Y']])])
    with pytest.raises(ValueError): run(source,out)
    import json
    assert json.loads((out/'report.json').read_text())['status']=='FAIL'


def test_unreadable_workbook_cannot_leave_success_report(tmp_path):
    source=create(tmp_path,[['A','Cedar','X']])
    out=tmp_path/'result'
    run(source,out)
    (source/'input.xlsx').write_bytes(b'not a workbook')
    with pytest.raises(ValueError,match='could not be read'): run(source,out)
    import json
    assert json.loads((out/'report.json').read_text())['status']=='FAIL'


def test_typo_and_ltd_limited_pair_requires_explicit_approval():
    records=[customer('old','Willow acount 003 Limited'),customer('keep','Willow account 003 Ltd')]
    pairs=propose(records)
    assert len(pairs)==1
    unchanged,_,errors=resolve(records,[])
    assert not errors and unchanged=={'old':'old','keep':'keep'}
    approved,_,errors=resolve(records,[('old','keep')])
    assert not errors and approved=={'old':'keep','keep':'keep'}
