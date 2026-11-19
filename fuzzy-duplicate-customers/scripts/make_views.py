import csv
from pathlib import Path
from customers import workbook, read_confirmed
from metadata import clean_properties
import openpyxl
from openpyxl.styles import Font

with Path('output/audit.csv').open() as f:
    audit={r['source_id']:r for r in csv.DictReader(f)}
with Path('output/proposed.csv').open() as f:
    proposals=list(csv.DictReader(f))
pair=next(p for p in proposals if {p['left_id'],p['right_id']}=={'C0003','C0403'})
assert audit['C0403']['retained_id']=='C0003'
approved=read_confirmed(Path('input/confirmed.csv'))
views=[
    ('Before',['Customer number','Name as entered'],[[key,audit[key]['original_name']] for key in ['C0003','C0403']]),
    ('After',['Number to keep','Name to keep'],[['C0003',audit['C0003']['original_name']]]),
    ('Proposed',['Customer numbers','Difference','Review'],[[pair['left_id']+' / '+pair['right_id'],'Typo; Ltd / Limited','Approval needed'],['C0041 / C0441','North / South','Not approved']]),
    ('Approved',['Old number','Number to keep','Decision'],[[s,t,'Approved'] for s,t in approved if s in ['C0402','C0403']])
]
for title,headers,rows in views:
    path=f'output/view-{title.lower()}.xlsx'
    workbook(path,[(title,headers,rows)])
    book=openpyxl.load_workbook(path)
    for sheet in book:
        for row in sheet:
            for cell in row:
                cell.font=Font(name='Carlito',size=20,bold=cell.row==1,color='FFFFFF' if cell.row==1 else '183D4A')
        for row in range(1,sheet.max_row+1):sheet.row_dimensions[row].height=38
        for col in sheet.column_dimensions.values():col.width*=1.6
    book.save(path)
    clean_properties(path)
