from pathlib import Path
import xlsxwriter, openpyxl
src=openpyxl.load_workbook('input/input.xlsx')
b=xlsxwriter.Workbook('macro-pack.xlsm'); b.set_properties({'author':'','last_modified_by':''}); b.add_vba_project('.build/vbaProject.bin'); b.set_vba_name('ThisWorkbook')
h=b.add_format({'bold':True,'bg_color':'#233F40','font_color':'white','font_size':13}); normal=b.add_format({'font_size':12})
for i,s in enumerate(src,1):
    ws=b.add_worksheet(s.title); ws.set_vba_name('Sheet'+str(i)); ws.set_column('A:C',25); ws.freeze_panes(1,0)
    for r,row in enumerate(s.values):
        for c,v in enumerate(row):
            if v is not None: ws.write(r,c,v,h if r==0 else normal)
groups=list(dict.fromkeys(str(r[1]) for r in list(src['Data'].values)[1:]))
for name in ['__MP_Split','__MP_Merged','__MP_Rename','__MP_Letters']:
    ws=b.add_worksheet(name); ws.set_column('A:C',24)
ws=b.add_worksheet('Template'); ws.write_string(0,0,Path('input/template.txt').read_text()); ws.set_column('A:A',75)
b.close()
