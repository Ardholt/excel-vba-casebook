"""Prepare paired scan crops and actual workbook values for the page figure."""
import csv,json,zipfile
from pathlib import Path
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import Font,Alignment,PatternFill
from openpyxl.drawing.image import Image as SheetImage
from extract import workbook
with Path('output/fields.csv').open() as f: fields=list(csv.DictReader(f))
batch=json.loads(Path('validation/batch.json').read_text())
faint=[x['image'] for x in batch['files'] if x['condition']=='hard to read' and x['affected_fields']==['Item code']]
focus=next(r for r in fields if r['image'] in faint and r['field']=='Item code' and not r['value'])
rows=[r for r in fields if r['image']==focus['image']]
report=json.loads(Path('output/report.json').read_text())
for name in ['Before','After']:
    values=[['Form number','Item code'],['',''],['Quantity','Amount'],['','']]
    if name=='After':
        for i,r in enumerate(rows):values[(i//2)*2+1][i%2]=r['value'] or ''
    p=Path('output')/('hero-'+name.lower()+'.xlsx')
    workbook(p,{name:values})
    wb=load_workbook(p);ws=wb[name]
    for col in 'AB':ws.column_dimensions[col].width=32
    for row in ws:
        header=row[0].row in (1,3)
        for c in row:
            c.font=Font(name='DejaVu Sans',size=20 if header else 28,color='FFFFFF' if header else '173D42',bold=header)
            c.fill=PatternFill('solid',fgColor='173D42' if header else 'FFFFFF')
            c.alignment=Alignment(vertical='center',wrap_text=True)
        ws.row_dimensions[row[0].row].height=28 if header else 44
    if name=='Before':
        with Image.open(Path('input/forms')/focus['image']) as im:
            for i,r in enumerate(rows):
                crop=Path('images')/f'hero-field-{i+1}.png'
                im.crop(tuple(map(int,r['box'].split(',')))).save(crop)
                pic=SheetImage(crop);pic.width=270;pic.height=34
                ws.add_image(pic,f"{'AB'[i%2]}{(i//2)*2+2}")
    else:ws['B2'].fill=PatternFill('solid',fgColor='F7E4B4')
    ws.auto_filter.ref=None;ws.print_area='A1:B4'
    ws.page_margins.left=0.2;ws.page_margins.right=0.2;ws.page_margins.top=0.2;ws.page_margins.bottom=0.2
    wb.save(p)
    with zipfile.ZipFile(p) as z: contents={n:z.read(n) for n in z.namelist()}
    contents['docProps/app.xml']=b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>'
    with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in contents.items():z.writestr(n,b)
Path('evidence/hero-source.json').write_text(json.dumps({'image':focus['image'],'field':focus['field'],'rectangle':focus['box'],'form_id':focus['form_id'],'value':focus['value'],'source':'output/fields.csv','fields':[{'field':r['field'],'rectangle':r['box'],'value':r['value'],'reason':r['reason']} for r in rows],'views':['output/hero-before.xlsx','output/hero-after.xlsx'],'presentation':'Actual scan field crops rearranged into a compact grid, paired with values from the reviewed workbook; blank cell shaded for checking.','forms':report['input_forms'],'remaining_checks':report['reviewed']['unresolved']},indent=2)+'\n')
