from pathlib import Path
import argparse, collections, csv, hashlib, json, re, time
from decimal import Decimal, ROUND_HALF_UP
from zipfile import ZipFile, ZIP_DEFLATED
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

PREFIX='__MP_'
REQUIRED={'Data':['record_id','group','value'],'MergeA':['record_id','group','value'],'MergeB':['record_id','group','value'],'RenameMap':['old','new'],'Recipients':['recipient_id','display_name','amount']}

def table(ws, required):
    for row in ws.iter_rows():
        vals=[str(c.value).strip() if c.value is not None else '' for c in row]
        if set(required)<=set(vals):
            heads=[v for v in vals if v]
            if len(heads)!=len(set(heads)): raise ValueError('Duplicate headers: '+ws.title)
            indexes=[vals.index(h) for h in required]
            rows=[]
            for rr in ws.iter_rows(min_row=row[0].row+1,values_only=True):
                if any(v is not None for v in rr): rows.append([rr[i] for i in indexes])
            return rows
    raise ValueError('Missing headers: '+ws.title)

def money(value):
    d=Decimal(str(value))
    if not d.is_finite(): raise ValueError('Non-finite amount')
    return format(d.quantize(Decimal('.01'),rounding=ROUND_HALF_UP),'.2f')

def valid_name(name):
    return bool(name) and len(name)<=31 and not re.search(r'[\[\]:*?/\\\x00-\x1f]',name) and not name.startswith("'") and not name.endswith("'") and name.casefold()!='history' and not name.casefold().startswith(PREFIX.casefold())

def rename_plan(names, rows):
    current={n.casefold():n for n in names}; seen=set(); out=[]
    protected=set(REQUIRED)|{'Instructions','Template','Config'}
    for old,new in rows:
        old,new=str(old or ''),str(new or '')
        key=old.casefold(); dest=new.casefold()
        if key in seen: status='rejected: repeated source'
        elif old.casefold() in {p.casefold() for p in protected} or old.casefold().startswith(PREFIX.casefold()): status='rejected: protected sheet'
        elif key not in current: status='rejected: missing source'
        elif not valid_name(new): status='rejected: invalid name'
        elif dest in current and dest!=key: status='rejected: destination exists'
        else:
            status='renamed'; del current[key]; current[dest]=new
        seen.add(key); out.append([old,new,status])
    return out

def render_letter(template,row):
    # One pass: inserted values are never interpreted as template tokens.
    vals=dict(zip(['recipient_id','display_name','amount'],[str(row[0]),str(row[1]),money(row[2])]))
    return re.sub(r'\{\{(recipient_id|display_name|amount)\}\}',lambda m:vals[m[1]],template)

def validate(book,template):
    tables={name:table(book[name],cols) for name,cols in REQUIRED.items()}
    for name in ['Data','MergeA','MergeB']:
        for row in tables[name]:
            if any(v is None for v in row): raise ValueError('Blank data field: '+name)
            if not Decimal(str(row[2])).is_finite(): raise ValueError('Invalid numeric value')
    unknown=re.sub(r'\{\{(recipient_id|display_name|amount)\}\}','',template)
    if '{{' in unknown or '}}' in unknown: raise ValueError('Unknown template token')
    template.encode('ascii'); ids=set()
    for rid,name,amt in tables['Recipients']:
        if not re.fullmatch('[A-Za-z0-9_-]+',str(rid or '')) or str(rid).lower() in ids or len(str(rid))>64 or str(rid).upper() in {'CON','PRN','AUX','NUL',*[f'COM{i}' for i in range(1,10)],*[f'LPT{i}' for i in range(1,10)]}: raise ValueError('Invalid or duplicate recipient ID')
        if not name: raise ValueError('Blank recipient name')
        str(name).encode('ascii'); ids.add(str(rid).lower()); money(amt)
        if Decimal(str(amt)) != Decimal(money(amt)): raise ValueError('Amount must use at most two decimal places')
    return tables

def style(ws):
    ws.freeze_panes='A2'; ws.sheet_view.showGridLines=False
    for c in ws[1]: c.font=Font(name='Calibri',size=13,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='233F40')
    for row in ws.iter_rows(min_row=2):
        for c in row: c.font=Font(name='Calibri',size=12); c.alignment=Alignment(vertical='top',wrap_text=True)
    for col,width in zip('ABCDE',[24,25,35,20,20]): ws.column_dimensions[col].width=width
    for i in range(1,min(ws.max_row,12)+1): ws.row_dimensions[i].height=27
    ws.print_options.horizontalCentered=True; ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A4; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1
    ws.print_area=f'A1:C{min(ws.max_row,8)}'
    ws.oddHeader.center.text=''; ws.oddFooter.center.text=''

def add(book,name,headers,rows):
    ws=book.create_sheet(name); ws.append(headers)
    for row in rows: ws.append(row)
    style(ws); return ws

def run(input_dir,output_dir):
    start=time.perf_counter(); inp=Path(input_dir).resolve(); out=Path(output_dir).resolve()
    if inp==out or inp in out.parents or out in inp.parents: raise ValueError('Input and output must be separate sibling folders')
    template=(inp/'template.txt').read_text('ascii'); book=openpyxl.load_workbook(inp/'input.xlsx'); tables=validate(book,template)
    if any(n.casefold().startswith(PREFIX.casefold()) for n in book.sheetnames): raise ValueError('Reserved output sheet prefix in input')
    out.mkdir(parents=True,exist_ok=True); letters=out/'letters'
    if letters.is_symlink(): raise ValueError('Letter directory must not be a symbolic link')
    letters.mkdir(exist_ok=True)
    for path in letters.iterdir():
        if path.is_file() and path.suffix=='.txt': path.unlink()
    groups=collections.OrderedDict()
    for row in tables['Data']: groups.setdefault(str(row[1]),[]).append(row)
    manifest=[]
    for i,(group,rows) in enumerate(groups.items(),1):
        name=f'{PREFIX}G{i:03}'; add(book,name,REQUIRED['Data'],rows); manifest.append([group,name,len(rows)])
    add(book,PREFIX+'Split',['group','sheet','rows'],manifest)
    merged=tables['MergeA']+tables['MergeB']; add(book,PREFIX+'Merged',REQUIRED['Data'],merged)
    renames=rename_plan(book.sheetnames,tables['RenameMap'])
    for old,new,status in renames:
        if status=='renamed': book[old].title=new
    add(book,PREFIX+'Rename',['old','new','status'],renames)
    letter_rows=[]
    for row in tables['Recipients']:
        content=render_letter(template,row); path=letters/(str(row[0])+'.txt'); path.write_text(content+'\n',encoding='ascii')
        letter_rows.append([row[0],money(row[2]),content])
    add(book,PREFIX+'Letters',['recipient_id','amount','letter'],letter_rows)
    with ZipFile(out/'letters.zip','w',ZIP_DEFLATED) as archive:
        for path in sorted(letters.glob('*.txt')): archive.write(path,path.name)
    summary=[['Split rows',len(tables['Data']),sum(len(r) for r in groups.values())],['Merged rows',len(merged),len(merged)],['Rename requests',len(renames),sum(r[2]=='renamed' for r in renames)],['Letters',len(tables['Recipients']),len(letter_rows)]]
    ws=add(book,'Results',['Action','Input','Output'],summary); book.active=book.index(ws)
    book.properties.creator=''; book.properties.lastModifiedBy=''; book.save(out/'after.xlsx')
    for name,heads,rows in [('split', ['group','sheet','rows'],manifest),('renames',['old','new','status'],renames),('merged',REQUIRED['Data'],merged)]:
        with (out/(name+'.csv')).open('w',newline='') as f: w=csv.writer(f); w.writerow(heads); w.writerows(rows)
    report={'implementation':'Python reference','input_records':{'split':len(tables['Data']),'merge_a':len(tables['MergeA']),'merge_b':len(tables['MergeB']),'rename':len(renames),'recipients':len(letter_rows)},'split':{'rows':len(tables['Data']),'groups':len(groups),'output_rows':sum(len(v) for v in groups.values()),'value_total':str(sum(Decimal(str(r[2])) for r in tables['Data']))},'merge':{'rows':len(merged),'duplicate_rows_preserved':sum(n-1 for n in collections.Counter(tuple(r) for r in merged).values()),'value_total':str(sum(Decimal(str(r[2])) for r in merged))},'rename':{'renamed':sum(r[2]=='renamed' for r in renames),'rejected':sum(r[2]!='renamed' for r in renames)},'letters':{'written':len(letter_rows)},'exclusions':{'split':0,'merge':0,'letters':0,'rename_reasons':renames},'runtime_seconds':round(time.perf_counter()-start,3),'checks':{'split_conserved':sum(len(r) for r in groups.values())==len(tables['Data']),'merge_conserved':len(merged)==len(tables['MergeA'])+len(tables['MergeB']),'letter_count_matches':len(letter_rows)==len(tables['Recipients'])}}
    (out/'report.json').write_text(json.dumps(report,indent=2)+'\n')
    (out/'run.txt').write_text(f"Python reference result\nSplit: {report['split']['rows']} rows into {len(groups)} groups; no lost rows.\nMerge: {len(merged)} rows; {report['merge']['duplicate_rows_preserved']} repeated rows retained.\nRename: {report['rename']['renamed']} renamed, {report['rename']['rejected']} rejected; see renames.csv.\nLetters: {len(letter_rows)} text files; nothing sent.\nElapsed: {report['runtime_seconds']} seconds.\n")
    return report

if __name__=='__main__':
    p=argparse.ArgumentParser(); p.add_argument('--input',default='input'); p.add_argument('--output',default='output'); a=p.parse_args(); print(json.dumps(run(a.input,a.output),indent=2))
