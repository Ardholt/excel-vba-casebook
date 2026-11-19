from pathlib import Path
import sys, shutil
from collections import Counter
import openpyxl, pytest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from reference import run, table, rename_plan, render_letter, valid_name
from check_output import check
ROOT=Path(__file__).resolve().parents[1]
@pytest.fixture
def work(tmp_path):
    inp=tmp_path/'input files'; out=tmp_path/'output files'; shutil.copytree(ROOT/'input',inp)
    return inp,out

def test_all_business_outputs(work):
    inp,out=work; report=run(inp,out); check(inp,out)
    assert report['split']['rows']==1200 and report['split']['groups']==6
    assert report['merge']['rows']==840 and report['merge']['duplicate_rows_preserved']==12
    assert report['rename']=={'renamed':2,'rejected':5}
    assert report['letters']['written']==80

def test_repeat_and_source_unchanged(work):
    inp,out=work; before=(inp/'input.xlsx').read_bytes(); first=run(inp,out); b1=openpyxl.load_workbook(out/'after.xlsx'); rows1={s.title:list(s.values) for s in b1}; letters1={p.name:p.read_bytes() for p in (out/'letters').glob('*.txt')}
    run(inp,out); b2=openpyxl.load_workbook(out/'after.xlsx'); assert rows1=={s.title:list(s.values) for s in b2}; assert letters1=={p.name:p.read_bytes() for p in (out/'letters').glob('*.txt')}; assert before==(inp/'input.xlsx').read_bytes()

def test_moved_headers_and_empty_groups(work):
    inp,out=work; b=openpyxl.load_workbook(inp/'input.xlsx')
    for name in ['Data','MergeA','MergeB','RenameMap','Recipients']: b[name].insert_rows(1,3)
    b['Data'].cell(5,2,'A/B'); b.save(inp/'input.xlsx'); run(inp,out); check(inp,out)

def test_empty_tables(work):
    inp,out=work; b=openpyxl.load_workbook(inp/'input.xlsx')
    for name in ['Data','MergeA','MergeB','Recipients']: b[name].delete_rows(2,b[name].max_row)
    b.save(inp/'input.xlsx'); r=run(inp,out); check(inp,out); assert r['merge']['rows']==r['split']['rows']==r['letters']['written']==0

@pytest.mark.parametrize('bad',['','Bad/name','X'*32,"'bad",'bad\x01','History','__MP_X'])
def test_invalid_sheet_name(bad): assert not valid_name(bad)

def test_sequential_collisions_and_case_insensitive():
    plan=rename_plan(['A','B','Keep'],[['A','B'],['B','A'],['Keep','keep'],['Keep','New']])
    assert [r[2] for r in plan]==['rejected: destination exists','rejected: destination exists','renamed','rejected: repeated source']

def test_letter_tokens_not_recursive():
    t='{{recipient_id}}|{{display_name}}|{{amount}}'
    assert render_letter(t,['R1','{{amount}}',12.5])=='R1|{{amount}}|12.50'

@pytest.mark.parametrize('kind',['duplicate_id','unsafe_id','unknown_token','missing_header','invalid_value'])
def test_bad_input_stops_before_output(work,kind):
    inp,out=work; b=openpyxl.load_workbook(inp/'input.xlsx')
    if kind=='duplicate_id': b['Recipients']['A3']='C001'
    if kind=='unsafe_id': b['Recipients']['A2']='../outside'
    if kind=='unknown_token': (inp/'template.txt').write_text('{{unknown}}')
    if kind=='missing_header': b['MergeB']['A1']='wrong'
    if kind=='invalid_value': b['Data']['C2']='not numeric'
    b.save(inp/'input.xlsx')
    with pytest.raises((ValueError,ArithmeticError)): run(inp,out)
    assert not out.exists()

def test_changed_inputs_not_fixed_answers(work):
    inp,out=work; b=openpyxl.load_workbook(inp/'input.xlsx'); b['Data'].append(['NEW','NEW GROUP',-7.5]); b['MergeB'].append([1.25,'NEW','G01']); b['Recipients'].append(['LAST','Last account',0]); b.save(inp/'input.xlsx')
    r=run(inp,out); check(inp,out); assert r['split']['rows']==1201 and r['merge']['rows']==841 and r['letters']['written']==81


def test_removed_recipient_clears_stale_letters(work):
    inp,out=work; run(inp,out); b=openpyxl.load_workbook(inp/'input.xlsx'); b['Recipients'].delete_rows(81); b.save(inp/'input.xlsx'); run(inp,out); check(inp,out); assert not (out/'letters/C080.txt').exists()

@pytest.mark.parametrize('rid',['CON','nul','LPT1','X'*65])
def test_portable_letter_filenames(work,rid):
    inp,out=work; b=openpyxl.load_workbook(inp/'input.xlsx'); b['Recipients']['A2']=rid; b.save(inp/'input.xlsx')
    with pytest.raises(ValueError): run(inp,out)
    assert not out.exists()
