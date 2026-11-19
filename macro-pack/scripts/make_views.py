from pathlib import Path
import openpyxl, json
from openpyxl.styles import Font, Alignment, PatternFill
from reference import add, style
b=openpyxl.Workbook(); b.remove(b.active)
for action,before,after in [('split','Data','__MP_Split'),('merge','MergeB','__MP_Merged'),('rename','RenameMap','__MP_Rename'),('letters','Recipients','__MP_Letters')]:
    for state,sheet in [('before',before),('after',after)]:
        source=openpyxl.load_workbook(f'output/vba/{action}-{state}.xlsx',data_only=True)
        rows=[list(r) for r in source[sheet].values]
        headings={'record_id':'Record','group':'Group','value':'Value','file':'Workbook','rows':'Rows','old':'Current tab','new':'Requested name','status':'Result','recipient_id':'Account','display_name':'Name','amount':'Amount'}
        rows[0]=[headings.get(v,v) for v in rows[0]]
        if action=='rename' and state=='after':
            for row in rows[1:]:
                status=row[2]
                if status=='renamed': row[2]='Changed'
                elif status=='rejected: missing source': row[2]='Sheet not found'
                elif status=='rejected: destination exists': row[2]='Name already used'
                elif status=='rejected: invalid name':
                    requested=str(row[1])
                    row[2]='Name too long' if len(requested)>31 else ('Starts with apostrophe' if requested.startswith("'") else ('Name contains /' if '/' in requested else 'Invalid tab name'))
        if action=='letters' and state=='after':
            ws=b.create_sheet('letters-after'); ws.append(['LETTER '+str(rows[1][0])]); ws.merge_cells('A1:C1'); ws.merge_cells('A2:C6'); ws['A2']=rows[1][2]
            style(ws); ws['A2'].font=Font(name='Calibri',size=17); ws['A2'].alignment=Alignment(wrap_text=True,vertical='center'); ws.print_area='A1:C6'
        else:
            ws=add(b,action+'-'+state,list(rows[0][:3]),[list(r[:3]) for r in rows[1:8]])
            if action=='rename':
                ws.column_dimensions['A'].width=19; ws.column_dimensions['B'].width=19; ws.column_dimensions['C'].width=30
                for row in ws.iter_rows(min_row=2):
                    for c in row: c.font=Font(name='Calibri',size=15)
                for i in range(2,9): ws.row_dimensions[i].height=58
                ws.print_area='A1:C8'
                ws['C1'].fill=PatternFill('solid',fgColor='233F40')
            else:
                ws.column_dimensions['A'].width=22; ws.column_dimensions['B'].width=22; ws.column_dimensions['C'].width=22
                for row in ws.iter_rows():
                    for c in row: c.font=Font(name='Calibri',size=16,bold=row[0].row==1,color='FFFFFF' if row[0].row==1 else '233F40')
                for i in range(1,9): ws.row_dimensions[i].height=31
                for row in ws.iter_rows(min_row=2):
                    for c in row:
                        c.alignment=Alignment(indent=1 if isinstance(c.value,str) else 0,horizontal='left' if isinstance(c.value,str) else 'right',vertical='center')
        ws.page_setup.orientation='landscape'; ws.sheet_view.showGridLines=False
b.properties.creator=''; b.properties.lastModifiedBy=''; b.save('evidence/views.xlsx')
Path('evidence/views.json').write_text(json.dumps({'source':'output/vba/<action>-<state>.xlsx','views':'Opening records from the recorded workbook snapshots; rename views show all seven requests; letter view shows the first complete letter. Column headings and rename decisions use plain wording. Names, amounts and row counts are unchanged; complete source logs are retained.'},indent=2)+'\n')

Path('evidence/views').mkdir(exist_ok=True)
for name in b.sheetnames:
    one=openpyxl.load_workbook('evidence/views.xlsx')
    for other in list(one.sheetnames):
        if other!=name: del one[other]
    one.save('evidence/views/'+name+'.xlsx')
