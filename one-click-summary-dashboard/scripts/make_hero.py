"""Compact proof graphic, typeset from saved button-run cells and event evidence."""
import hashlib
import json
from decimal import Decimal
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import openpyxl

source=Path('output/vba-after.xlsx')
book=openpyxl.load_workbook(source,data_only=True)
d=book['Dashboard']
record=json.loads(Path('output/vba-run.json').read_text())
assert record['result']=='PASS' and all(r['button_activated'] for r in record['runs'])
before,after,repeat=record['runs']
assert Decimal(str(d['A3'].value))==Decimal(after['paid_revenue'])
months=[(d.cell(r,1).value,d.cell(r,2).value) for r in range(12,12+len(after['paid_by_month']))]
start=14+len(months)
categories=[(d.cell(r,1).value,d.cell(r,2).value) for r in range(start,start+len(after['paid_by_category']))]
assert dict((k,f'{v:.2f}') for k,v in months)==after['paid_by_month']
assert dict((k,f'{v:.2f}') for k,v in categories)==after['paid_by_category']
navy,teal,muted,light='#142F43','#087F73','#486171','#E8F3F1'
image=Image.new('RGB',(760,724),'white');draw=ImageDraw.Draw(image)
def font(n,bold=False):
    return ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans'+('-Bold' if bold else '')+'.ttf',n)
def text(x,y,s,n=26,fill=navy,bold=False,anchor=None):
    draw.text((x,y),s,font=font(n,bold),fill=fill,anchor=anchor)
def money(v):return f'£{Decimal(str(v)):,.2f}'
text(20,12,'PAID SALES · ALL MONTHS & CATEGORIES',24,bold=True)
text(18,46,money(after['paid_revenue']),58,teal,True)
text(20,118,f"Was {money(before['paid_revenue'])}  ·  +{money(Decimal(after['paid_revenue'])-Decimal(before['paid_revenue']))}",27)
draw.rounded_rectangle((20,161,273,207),radius=7,fill=teal)
text(36,168,record['button']['caption'],28,'white',True)
text(288,169,'→ new orders counted',28)
for x,title in [(20,'Month · GBP'),(400,'Category · GBP')]:
    draw.rectangle((x,227,x+340,268),fill=navy);text(x+10,232,title,26,'white',True)
for x,items in [(20,months),(400,categories)]:
    for i,(label,value) in enumerate(items):
        y=275+i*30
        if label in ('2026-01','Repairs'):draw.rectangle((x,y-1,x+340,y+29),fill=light)
        text(x+5,y,label,25,bold=label in ('2026-01','Repairs'))
        text(x+335,y,f'{value:,.2f}',25,anchor='ra')
        draw.line((x,y+29,x+340,y+29),fill='#DAE2E6')
text(405,519,'Still owed · separate',24,muted)
text(405,553,money(after['pending_amount']),32,teal,True)
text(405,606,f"{after['rejected_records']} rows for review",25)
draw.rectangle((20,679,740,722),fill=light)
text(32,684,'Repeat press: '+money(repeat['paid_revenue']),29,teal,True)
image.save('media/hero.png')
book.close()
Path('evidence/hero.json').write_text(json.dumps({
    'file':'media/hero.png','width':760,'height':724,
    'source':str(source),'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),
    'button_run':'output/vba-run.json','regions':['Dashboard!A3','Dashboard!A12:B24','Dashboard!A27:B33'],
    'visible_months':[x[0] for x in months],'categories':[x[0] for x in categories],
    'presentation':'Typeset proof graphic from saved workbook cells and recorded button actions; not an application screenshot.',
    'sha256':hashlib.sha256(Path('media/hero.png').read_bytes()).hexdigest()},indent=2)+'\n')
print('PASS: hero shows every month, category, overall total and recorded button result')
