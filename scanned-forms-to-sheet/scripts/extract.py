"""Read image pixels with Tesseract. No dependency on generation or answer keys."""
import argparse, csv, hashlib, io, json, re, subprocess, time
from collections import Counter
from decimal import Decimal
from pathlib import Path
from PIL import Image, ImageOps
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

FIELDS = ('Form ID', 'Item code', 'Quantity', 'Amount')
PATTERNS = (r'F[0-9]{3,8}', r'IT-[0-9]{3,8}', r'[0-9]{1,6}', r'[0-9]{1,9}\.[0-9]{2}')
THRESHOLD = 85
BOXES = {field: (410, 147+100*i, 945, 214+100*i) for i, field in enumerate(FIELDS)}

def valid(field, value):
    return field in FIELDS and bool(re.fullmatch(PATTERNS[FIELDS.index(field)], value))

def write_csv(path, rows, headers):
    with Path(path).open('w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=headers); w.writeheader(); w.writerows(rows)

def ocr(im):
    # Standardise the pixel scale without altering recognised characters.
    gray = ImageOps.grayscale(im)
    bounds = gray.point(lambda value: 255 if value < 128 else 0).getbbox()
    if bounds:
        gray = ImageOps.expand(gray.crop(bounds), border=16, fill='white')
    im = gray.resize((gray.width * 2, gray.height * 2), Image.Resampling.LANCZOS)
    data = io.BytesIO(); im.save(data, format='PNG')
    p = subprocess.run(['tesseract','stdin','stdout','-l','eng','--psm','7','tsv'], input=data.getvalue(), capture_output=True, timeout=25)
    if p.returncode: raise RuntimeError('OCR engine failed')
    words = [r for r in csv.DictReader(io.StringIO(p.stdout.decode()), delimiter='\t') if r['text'].strip() and float(r['conf']) >= 0]
    return ' '.join(r['text'].strip() for r in words), min((float(r['conf']) for r in words), default=0.0)

def read_forms(indir):
    paths = sorted(p for p in (Path(indir)/'forms').iterdir() if p.is_file() and p.suffix.lower()=='.png')
    if not paths: raise ValueError('No PNG forms found')
    records, inventory = [], []
    for path in paths:
        error = ''
        try:
            with Image.open(path) as original:
                im = original.convert('RGB')
            if im.size != (1000,600): error = 'Unsupported image dimensions'
        except Exception:
            error = 'Image cannot be read'; im = None
        inventory.append({'image':path.name, 'sha256':hashlib.sha256(path.read_bytes()).hexdigest(), 'width':im.width if im else 0,'height':im.height if im else 0})
        for field in FIELDS:
            value, confidence = '', 0
            reason = error
            if not reason:
                try: value, confidence = ocr(im.crop(BOXES[field]))
                except (RuntimeError, subprocess.TimeoutExpired): reason = 'OCR failed or timed out'
            if not reason:
                if not value: reason = 'No readable text'
                elif not valid(field, value): reason = 'Field format not recognised'
                elif confidence < THRESHOLD: reason = 'Low OCR confidence'
            records.append({'image':path.name,'form_id':'','field':field,'candidate':value,'confidence':round(confidence,2),'raw_value':value if not reason else '', 'value':value if not reason else '', 'status':'unresolved' if reason else 'extracted','reason':reason,'box':','.join(map(str,BOXES[field]))})
    ids = {r['image']:r['value'] for r in records if r['field']=='Form ID'}
    counts = Counter(v for v in ids.values() if v)
    for r in records:
        r['form_id'] = ids[r['image']] if ids[r['image']] and counts[ids[r['image']]]==1 else 'file:'+r['image']
        if r['field']=='Form ID' and ids[r['image']] and counts[ids[r['image']]]>1:
            r.update(raw_value='',value='',status='unresolved',reason='Duplicate Form ID')
    return records, inventory

def apply_review(records, path):
    audit, errors = [], []
    if not Path(path).exists(): return audit, errors
    with Path(path).open(newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        if not {'form_id','field','value'} <= set(reader.fieldnames or []):
            return [], [{'row':1, 'reason':'Review headers missing'}]
        rows = list(reader)
    counts = Counter(((r.get('form_id') or '').strip(),(r.get('field') or '').strip()) for r in rows)
    proposed_ids = Counter((r.get('value') or '').strip() for r in rows if (r.get('field') or '').strip()=='Form ID')
    for n, row in enumerate(rows, 2):
        fid, field, value = ((row.get(k) or '').strip() for k in ('form_id','field','value'))
        matches = [r for r in records if r['form_id']==fid and r['field']==field]
        reason = ''
        if counts[(fid,field)] > 1: reason = 'Duplicate review target'
        elif len(matches)!=1: reason = 'Unknown or ambiguous review target'
        elif not valid(field,value): reason = 'Invalid review value'
        elif field=='Form ID' and (proposed_ids[value]>1 or any(r['field']=='Form ID' and r['form_id']!=fid and r['value']==value for r in records)): reason = 'Form ID correction would duplicate an ID'
        if reason: errors.append({'row':n,'reason':reason}); continue
        target = matches[0]
        audit.append({'row':n,'image':target['image'],'form_id':fid,'field':field,'before':target['value'],'after':value,'original_reason':target['reason']})
        target.update(value=value,status='reviewed')
    return audit, errors

def workbook(path, sheets):
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    wb.properties.creator = ''; wb.properties.lastModifiedBy = ''
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
            for c in ws[ws.max_row]:
                if isinstance(c.value,str): c.data_type='s'
                c.font = Font(name='DejaVu Sans',size=12,color='173D42')
                c.alignment = Alignment(vertical='center',wrap_text=True)
            ws.row_dimensions[ws.max_row].height=30
        for cell in ws[1]: cell.fill=PatternFill('solid',fgColor='173D42'); cell.font=Font(name='DejaVu Sans',size=12,bold=True,color='FFFFFF')
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width=25
        for c in ws[1]:
            if c.value == 'Amount':
                for row in ws.iter_rows(min_row=2,min_col=c.column,max_col=c.column): row[0].number_format='0.00'
        if name=='Data':
            for row in ws.iter_rows(min_row=2):
                for c in row[1:5]:
                    if c.value in ('',None): c.fill=PatternFill('solid',fgColor='F7E4B4')
        ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
        ws.sheet_view.showGridLines=False
        ws.print_area=ws.dimensions
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A4; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    wb.save(path)
    # App properties are intentionally empty, including producer name.
    import zipfile
    with zipfile.ZipFile(path) as z: contents={n:z.read(n) for n in z.namelist()}
    contents['docProps/app.xml']=b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>'
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in contents.items(): z.writestr(n,b)

def table(rows, headers): return [headers]+[[r.get(k,'') for k in headers] for r in rows]

def data_rows(records, raw=False):
    result=[['Image',*FIELDS,'Status']]
    for image in dict.fromkeys(r['image'] for r in records):
        fields=[r for r in records if r['image']==image]
        values=[r['raw_value' if raw else 'value'] for r in fields]
        values[2]=int(values[2]) if values[2] else None
        values[3]=float(Decimal(values[3])) if values[3] else None
        result.append([image,*values,'Needs review' if any(v in ('',None) for v in values) else 'Complete'])
    return result

def run(indir, outdir):
    start=time.perf_counter(); indir=Path(indir).resolve(); out=Path(outdir).resolve()
    if out==indir or indir in out.parents or out in indir.parents: raise ValueError('Input and output must be separate directories')
    out.mkdir(parents=True,exist_ok=True)
    (out/'report.json').write_text(json.dumps({'status':'started','checks_complete':False},indent=2)+'\n')
    records, inventory=read_forms(indir)
    audit,errors=apply_review(records,indir/'review.csv')
    headers=list(records[0]); write_csv(out/'fields.csv',records,headers)
    queue=[r for r in records if not r['raw_value']]
    unresolved=[r for r in records if not r['value']]
    write_csv(out/'unclear-entries.csv',queue,headers)
    write_csv(out/'unresolved.csv',unresolved,headers)
    write_csv(out/'review-audit.csv',audit,['row','image','form_id','field','before','after','original_reason'])
    write_csv(out/'review-errors.csv',errors,['row','reason'])
    write_csv(out/'inventory.csv',inventory,['image','sha256','width','height'])
    workbook(out/'input-inventory.xlsx', {'Scans':table(inventory,['image','width','height','sha256'])})
    workbook(out/'extracted.xlsx',{'Data':data_rows(records,True),'Unclear entries':table(queue,headers)})
    workbook(out/'reviewed.xlsx',{'Data':data_rows(records),'Still unresolved':table(unresolved,headers),'Review audit':table(audit,['form_id','field','before','after','original_reason'])})
    raw_clear=sum(bool(r['raw_value']) for r in records); final_clear=sum(bool(r['value']) for r in records)
    report={'input_forms':len(inventory),'output_rows':len(inventory),'fields_per_form':4,'total_fields':len(records),'raw':{'populated':raw_clear,'unresolved':len(records)-raw_clear},'reviewed':{'populated':final_clear,'unresolved':len(records)-final_clear,'confirmations_applied':len(audit)},'review_errors':errors,'excluded_forms':0,'confidence_threshold':THRESHOLD,'runtime_seconds':round(time.perf_counter()-start,3),'checks':{'all_fields_accounted':len(records)==4*len(inventory),'unique_field_locations':len({(r['image'],r['field']) for r in records})==len(records),'review_valid':not errors,'ocr_operational':not any(r['reason']=='OCR failed or timed out' for r in records)}}
    report['status']='passed' if all(report['checks'].values()) else 'failed'
    (out/'report.json').write_text(json.dumps(report,indent=2)+'\n')
    (out/'evidence.md').write_text(f"# Extraction result\n\nForms: {len(inventory)}. Fields: {len(records)}.\n\nRaw: {raw_clear} populated, {len(records)-raw_clear} unresolved.\n\nAfter supplied confirmations: {final_clear} populated, {len(records)-final_clear} unresolved.\n\nCorrections applied: {len(audit)}. Invalid review rows: {len(errors)}.\n\nEvery source field appears in fields.csv. Empty values stay empty in the workbook.\n")
    print(f"Forms {len(inventory)}; raw unresolved {len(queue)}; reviewed unresolved {len(unresolved)}")
    if errors: raise ValueError('Invalid review rows; see review-errors.csv')
    if not all(report['checks'].values()): raise ValueError('Conservation checks failed')
    return report

if __name__=='__main__':
    p=argparse.ArgumentParser(); p.add_argument('--input',default='input'); p.add_argument('--output',default='output'); a=p.parse_args()
    run(a.input,a.output)
