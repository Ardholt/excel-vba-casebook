"""Build a narrow evidence sheet from the actual extraction and accuracy report."""
import csv,json
from pathlib import Path
from PIL import Image
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as SheetImage
from openpyxl.styles import Font,PatternFill,Alignment
from extract import workbook
Path('images').mkdir(parents=True,exist_ok=True)
out=Path('output')
report=json.loads((out/'report.json').read_text()); score=json.loads((out/'accuracy.json').read_text())
with (out/'fields.csv').open() as f: fields=list(csv.DictReader(f))
with (out/'review-audit.csv').open() as f: audit=list(csv.DictReader(f))
focus=next(r for r in fields if not r['raw_value'] and r['status']=='reviewed')
fix=next(r for r in audit if r['image']==focus['image'] and r['field']==focus['field'])
box=tuple(map(int,focus['box'].split(',')))
with Image.open(Path('input/forms')/focus['image']) as im:
    crop=im.crop(box)
    crop.save('images/unclear-field.png')
workbook(out/'proof.xlsx',{'Proof':[['UNCLEAR ON THE SCAN', '', 'BLANK IN THE SHEET',''],
    ['Source: '+focus['form_id'],'','Raw extraction',''],
    ['Amount','','Form ID',focus['form_id']],
    ['','','Amount','Left blank'],
    ['','','Status','Needs review'],
    ['Review file: '+fix['after']+' | raw reading kept','','',''],
    [f"Raw: {score['raw']['correct']}/{score['denominator']} correct | {score['raw']['unresolved']} unresolved | {score['raw']['errors']} errors",'','',''],
    [f"Reviewed: {score['reviewed']['correct']}/{score['denominator']} correct | {score['reviewed']['unresolved']} unresolved | {score['reviewed']['errors']} errors",'','','']]})
wb=load_workbook(out/'proof.xlsx');ws=wb['Proof']
for col in 'ABCD':ws.column_dimensions[col].width=18
for r in range(1,9):
    ws.row_dimensions[r].height=38
    for c in ws[r]: c.font=Font(name='DejaVu Sans',size=16,color='173D42'); c.alignment=Alignment(vertical='center',wrap_text=True)
for r in (1,2):
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
for r in (6,7,8): ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=4)
for c in ws[1]:c.font=Font(name='DejaVu Sans',size=15,bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='173D42')
for c in ws[6]:c.fill=PatternFill('solid',fgColor='F7E4B4')
ws['D4'].fill=PatternFill('solid',fgColor='F7E4B4')
for c in ws[8]:c.font=Font(name='DejaVu Sans',size=14,bold=True,color='173D42')
pic=SheetImage('images/unclear-field.png');pic.width=240;pic.height=30;ws.add_image(pic,'A4')
ws.print_options.horizontalCentered=True
ws.page_margins.left=0.2;ws.page_margins.right=0.2;ws.page_margins.top=0.2;ws.page_margins.bottom=0.2
ws.auto_filter.ref=None;ws.print_area='A1:D8'
wb.save(out/'proof.xlsx')
# Remove application attribution after the presentation edit.
import zipfile
p=out/'proof.xlsx'
with zipfile.ZipFile(p) as z: contents={n:z.read(n) for n in z.namelist()}
contents['docProps/app.xml']=b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>'
with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
    for n,b in contents.items():z.writestr(n,b)

# A second, legible comparison uses actual pixel crops and matching extracted cells.
from extract import FIELDS, BOXES
clear_image=next(name for name in dict.fromkeys(r['image'] for r in fields) if all(r['raw_value'] for r in fields if r['image']==name))
clear=[r for r in fields if r['image']==clear_image]
rows=[['SCAN PIXELS','','EXTRACTED ROW','']]+[['','',r['field'],r['raw_value']] for r in clear]
workbook(out/'readable-proof.xlsx',{'Readable':rows})
wb=load_workbook(out/'readable-proof.xlsx');ws=wb['Readable']
for col in 'ABCD':ws.column_dimensions[col].width=18
ws.merge_cells('A1:B1');ws.merge_cells('C1:D1')
for r in range(1,6):
    ws.row_dimensions[r].height=60 if r>1 else 40
    for c in ws[r]: c.font=Font(name='DejaVu Sans',size=18,color='173D42');c.alignment=Alignment(vertical='center',wrap_text=True)
for c in ws[1]:c.font=Font(name='DejaVu Sans',size=15,bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='173D42')
with Image.open(Path('input/forms')/clear_image) as im:
    for i,r in enumerate(clear,2):
        crop_path=Path('images')/('readable-'+str(i-1)+'.png')
        im.crop(BOXES[r['field']]).save(crop_path)
        pic=SheetImage(crop_path);pic.width=240;pic.height=30;ws.add_image(pic,'A'+str(i))
ws.auto_filter.ref=None;ws.print_area='A1:D5'
ws.page_margins.left=0.2;ws.page_margins.right=0.2;ws.page_margins.top=0.2;ws.page_margins.bottom=0.2
wb.save(out/'readable-proof.xlsx')
p=out/'readable-proof.xlsx'
with zipfile.ZipFile(p) as z: contents={n:z.read(n) for n in z.namelist()}
contents['docProps/app.xml']=b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>'
with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
    for n,b in contents.items():z.writestr(n,b)
Path('evidence/proof-sources.json').write_text(json.dumps({'unclear':{'image':focus['image'],'field':focus['field'],'box':list(box)},'readable':{'image':clear_image,'fields':[{'field':r['field'],'box':r['box'],'value':r['raw_value']} for r in clear]},'reports':['output/fields.csv','output/accuracy.json','output/review-audit.csv']},indent=2)+'\n')
