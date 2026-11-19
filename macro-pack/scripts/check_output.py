import sys
from pathlib import Path
from collections import Counter
from zipfile import ZipFile, is_zipfile
import openpyxl
from reference import validate,table,REQUIRED,rename_plan,render_letter

def check(inp,out):
    inp,out=Path(inp),Path(out); source=openpyxl.load_workbook(inp/'input.xlsx'); result=openpyxl.load_workbook(out/'after.xlsx'); t=(inp/'template.txt').read_text(); tables=validate(source,t)
    manifest=list(result['__MP_Split'].values)[1:]; split=[]
    assert not any(n.startswith('__MP_G') for n in result.sheetnames)
    assert {p.name for p in (out/'split').glob('group-*.xlsx')}=={r[1] for r in manifest}
    for group,name,count in manifest:
        path=out/'split'/name
        assert is_zipfile(path), 'Split output must be real XLSX'
        separate=openpyxl.load_workbook(path)
        assert separate.sheetnames==['Data'], 'Each workbook contains only its group'
        rows=table(separate['Data'],REQUIRED['Data']); assert len(rows)==count; assert all(str(r[1])==str(group) for r in rows); split+=rows
    with ZipFile(out/'split-workbooks.zip') as archive:
        assert set(archive.namelist())=={r[1] for r in manifest}
        for name in archive.namelist(): assert archive.read(name)==(out/'split'/name).read_bytes()
    assert Counter(map(tuple,split))==Counter(map(tuple,tables['Data']))
    assert table(result['__MP_Merged'],REQUIRED['Data'])==tables['MergeA']+tables['MergeB']
    plan=rename_plan(source.sheetnames,tables['RenameMap']); assert [list(r) for r in list(result['__MP_Rename'].values)[1:]]==plan
    for old,new,status in plan:
        if status=='renamed': assert list(result[new].values)==list(source[old].values)
        elif old in source.sheetnames: assert list(result[old].values)==list(source[old].values)
    expected={str(r[0])+'.txt':render_letter(t,r)+'\n' for r in tables['Recipients']}
    actual={p.name:p.read_text() for p in (out/'letters').glob('*.txt')}
    assert actual==expected
    with ZipFile(out/'letters.zip') as archive:
        assert {n:archive.read(n).decode('ascii') for n in archive.namelist()}==expected
    print('CHECK split rows and values conserved; merge aligned; rejected sheets preserved; letters isolated')
if __name__=='__main__': check(*sys.argv[1:])
