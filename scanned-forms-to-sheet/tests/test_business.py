import csv,json,sys
from pathlib import Path
from collections import Counter
import pytest
from PIL import Image
import openpyxl
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from extract import read_forms,apply_review,run,valid,FIELDS,workbook
from evaluate import evaluate
from make_forms import draw_form

@pytest.fixture
def fresh(tmp_path):
    source=tmp_path/'new input'; (source/'forms').mkdir(parents=True)
    draw_form(source/'forms'/'unexpected name.png',['F987','IT-853','27','415.29'])
    return source

@pytest.fixture
def entries():
    return [{'image':'scan.png','form_id':'F987','field':f,'candidate':'','confidence':0,'raw_value':'','value':'','status':'unresolved','reason':'No readable text','box':'0,0,1,1'} for f in FIELDS]

def reviews(tmp_path, rows):
    path=tmp_path/'review.csv'
    with path.open('w',newline='') as f:
        w=csv.writer(f);w.writerow(['value','field','form_id']);w.writerows(rows)
    return path

def test_fresh_pixels_not_filename_or_key(fresh,tmp_path):
    report=run(fresh,tmp_path/'new output')
    rows=list(openpyxl.load_workbook(tmp_path/'new output/reviewed.xlsx',data_only=True)['Data'].values)
    expected=('F987','IT-853',27,415.29)
    actual=rows[1][1:5]
    assert sum(value is not None for value in actual)>=3
    assert all(value is None or value==answer for value,answer in zip(actual,expected))
    with (tmp_path/'new output/fields.csv').open() as f:
        fields=list(csv.DictReader(f))
    assert {r['field'] for r in fields}==set(FIELDS)
    assert all(r['value'] or (r['status']=='unresolved' and r['reason']) for r in fields)
    assert report['total_fields']==4
    assert report['raw']['unresolved']==sum(value is None for value in actual)

def test_obscured_field_remains_blank(fresh):
    draw_form(fresh/'forms'/'unexpected name.png',['F987','IT-853','27','415.29'],['Amount'])
    records,_=read_forms(fresh)
    r=next(r for r in records if r['field']=='Amount')
    assert r['value']=='' and r['status']=='unresolved' and r['reason']
    assert r['box']=='410,447,945,514'

def test_unsupported_resolution_accounts_for_every_field(fresh):
    Image.new('RGB',(100,100),'white').save(fresh/'forms'/'unexpected name.png')
    records,_=read_forms(fresh)
    assert len(records)==4 and all(r['status']=='unresolved' for r in records)

def test_corrupt_png_accounts_for_every_field(fresh):
    (fresh/'forms'/'unexpected name.png').write_bytes(b'not an image')
    records,_=read_forms(fresh)
    assert len(records)==4 and all(r['reason']=='Image cannot be read' for r in records)

def test_duplicate_form_ids_are_not_review_targets(fresh):
    draw_form(fresh/'forms'/'second.png',['F987','IT-723','31','78.50'])
    records,_=read_forms(fresh)
    ids=[r for r in records if r['field']=='Form ID']
    assert len(ids)==2 and all(r['reason']=='Duplicate Form ID' and not r['value'] for r in ids)
    assert len({r['form_id'] for r in ids})==2

def test_review_audits_original_and_preserves_raw(entries,tmp_path):
    audit,errors=apply_review(entries,reviews(tmp_path,[['163.52','Amount','F987']]))
    assert not errors and len(audit)==1 and audit[0]['before']=='' and audit[0]['after']=='163.52'
    assert entries[3]['value']=='163.52' and entries[3]['raw_value']==''

def test_invalid_review_is_reported(entries,tmp_path):
    audit,errors=apply_review(entries,reviews(tmp_path,[['ten','Quantity','F987']]))
    assert not audit and len(errors)==1 and not entries[2]['value']

def test_unknown_review_is_reported(entries,tmp_path):
    audit,errors=apply_review(entries,reviews(tmp_path,[['21','Quantity','F111']]))
    assert not audit and len(errors)==1

def test_conflicting_reviews_do_not_apply_either(entries,tmp_path):
    audit,errors=apply_review(entries,reviews(tmp_path,[['21','Quantity','F987'],['22','Quantity','F987']]))
    assert not audit and len(errors)==2 and not entries[2]['value']

def test_no_review_leaves_unresolved(entries,tmp_path):
    audit,errors=apply_review(entries,tmp_path/'absent.csv')
    assert not audit and not errors and all(not r['value'] for r in entries)

def test_repeat_run_is_same_business_result(fresh,tmp_path):
    out=tmp_path/'result'
    run(fresh,out); first=(out/'fields.csv').read_bytes()
    run(fresh,out)
    assert (out/'fields.csv').read_bytes()==first
    assert len(list(openpyxl.load_workbook(out/'reviewed.xlsx')['Data'].values))==2

def test_input_output_overlap_refused(fresh):
    with pytest.raises(ValueError,match='separate'): run(fresh,fresh/'output')

def test_empty_input_refused(tmp_path):
    (tmp_path/'forms').mkdir()
    with pytest.raises(ValueError,match='No PNG'): read_forms(tmp_path)

def test_accuracy_counts_wrong_values_and_unresolved(tmp_path):
    out=tmp_path/'score';out.mkdir()
    with (out/'fields.csv').open('w',newline='') as f:
        w=csv.writer(f);w.writerow(['image','field','raw_value','value']);w.writerows([['x.png','Quantity','','21'],['x.png','Amount','9.00','8.00']])
    key=tmp_path/'key.csv';key.write_text('image,field,value\nx.png,Quantity,21\nx.png,Amount,8.00\n')
    report=evaluate(out,key)
    assert report['denominator']==2
    assert report['raw']=={'correct':0,'unresolved':1,'errors':1}
    assert report['reviewed']=={'correct':2,'unresolved':0,'errors':0}

def test_accuracy_refuses_partial_key(tmp_path):
    (tmp_path/'fields.csv').write_text('image,field,raw_value,value\nx.png,Amount,1.00,1.00\n')
    (tmp_path/'key.csv').write_text('image,field,value\ny.png,Amount,1.00\n')
    with pytest.raises(ValueError,match='exactly once'): evaluate(tmp_path,tmp_path/'key.csv')

def test_low_confidence_is_never_accepted(fresh,monkeypatch):
    monkeypatch.setattr('extract.ocr',lambda im:('F987',42))
    records,_=read_forms(fresh)
    assert not records[0]['value'] and records[0]['reason']=='Low OCR confidence'

def test_text_cannot_become_workbook_formula(tmp_path):
    p=tmp_path/'book.xlsx';workbook(p,{'Data':[['Image'],['=2+2']]})
    assert openpyxl.load_workbook(p)['Data']['A2'].data_type=='s'

@pytest.mark.parametrize('field,value',[('Amount','12,50'),('Amount','-12.50'),('Quantity','2.5'),('Form ID','F98O')])
def test_unsupported_values_rejected(field,value):
    assert not valid(field,value)

def test_whitespace_does_not_hide_duplicate_review(entries,tmp_path):
    audit,errors=apply_review(entries,reviews(tmp_path,[['21','Quantity','F987'],['22',' Quantity ',' F987 ']]))
    assert not audit and len(errors)==2

def test_form_id_review_cannot_create_duplicate(entries,tmp_path):
    entries.append(dict(entries[0],image='another.png',form_id='F123',value='F123'))
    audit,errors=apply_review(entries,reviews(tmp_path,[['F123','Form ID','F987']]))
    assert not audit and errors[0]['reason']=='Form ID correction would duplicate an ID'

def test_ocr_operational_failure_cannot_report_success(fresh,tmp_path,monkeypatch):
    def broken(im): raise RuntimeError('OCR failed')
    monkeypatch.setattr('extract.ocr',broken)
    out=tmp_path/'failed output'
    with pytest.raises(ValueError,match='Conservation checks failed'): run(fresh,out)
    report=json.loads((out/'report.json').read_text())
    assert report['total_fields']==4 and not report['checks']['ocr_operational']
