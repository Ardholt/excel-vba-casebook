from pathlib import Path
from decimal import Decimal
import openpyxl
from reference import add,style

def build():
    b=openpyxl.Workbook(); b.remove(b.active)
    add(b,'Instructions',['Task','Before','Ready after run'],[['Split','1,200 mixed rows','Group sheets'],['Merge','Columns in different order','One aligned table'],['Rename','Conflicts in the map','Accepted + rejected'],['Letters','80 records + a template','One file per record']])
    rows=[[f'R{i:04}',f'G{(i-1)%6+1:02}',float(Decimal(i%97+1)*Decimal('2.75'))] for i in range(1,1201)]
    add(b,'Data',['record_id','group','value'],rows)
    a=rows[:480]; other=rows[480:828]+rows[:12]
    add(b,'MergeA',['record_id','group','value'],a)
    add(b,'MergeB',['value','record_id','group'],[[r[2],r[0],r[1]] for r in other])
    maps=[['North old','North'],['South old','South'],['Missing','New tab'],['Keep A','North'],['Keep B','Bad/name'],['Keep C','X'*32],['Keep D',"'Bad"]]
    add(b,'RenameMap',['old','new'],maps)
    add(b,'Recipients',['recipient_id','display_name','amount'],[[f'C{i:03}',f'Account {i:03}',i*12.25] for i in range(1,81)])
    for name in ['North old','South old','Keep A','Keep B','Keep C','Keep D']: add(b,name,['record_id','group','value'],[[name,'Retain this content',17]])
    template='Reference: {{recipient_id}}\nDear {{display_name}},\nYour statement amount is GBP {{amount}}.\nPlease keep this letter with your records.'
    Path('input/template.txt').write_text(template,encoding='ascii')
    b.properties.creator=''; b.properties.lastModifiedBy=''; b.save('input/input.xlsx')
if __name__=='__main__': build()
