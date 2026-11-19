"""Execute and measure the same pipeline inside the shared container."""
import argparse,json,subprocess,time,platform,re,os,hashlib
from datetime import datetime,timezone
from pathlib import Path
import openpyxl,PIL,pytest
from extract import run
from evaluate import evaluate
p=argparse.ArgumentParser();p.add_argument('--input',default='input');p.add_argument('--output',default='output');p.add_argument('--benchmark',action='store_true');a=p.parse_args()
start=time.perf_counter()
def source_hashes():
    return {str(p.relative_to(a.input)):hashlib.sha256(p.read_bytes()).hexdigest() for p in Path(a.input).rglob('*') if p.is_file()}
before=source_hashes()
run(a.input,a.output)
if a.benchmark:
    accuracy=evaluate(a.output,'validation/answer-key.csv')
    assert accuracy['raw']['errors']==0 and accuracy['reviewed']['errors']==0, 'Accepted value differs from answer key'
else:
    Path(a.output,'accuracy.json').write_text(json.dumps({'status':'not assessed','reason':'Replacement batch; no answer key supplied.'},indent=2)+'\n')
    Path(a.output,'accuracy-fields.csv').write_text('phase,image,field,expected,actual,result\n')
subprocess.run(['python3','scripts/check_output.py',a.output],check=True)
test_start=time.perf_counter()
proc=subprocess.run(['python3','-m','pytest','-q','-p','no:cacheprovider','tests'],text=True,capture_output=True)
print(proc.stdout)
summary=[s for s in proc.stdout.splitlines() if re.search(r'\d+ (passed|failed|error)',s)]
versions={'Python':platform.python_version(),'openpyxl':openpyxl.__version__,'Pillow':PIL.__version__,'pytest':pytest.__version__,'Tesseract':subprocess.check_output(['tesseract','--version'],text=True).splitlines()[0],'LibreOffice':subprocess.check_output(['libreoffice','--version'],text=True).strip()}
assert before==source_hashes(), 'Source input changed during the run'
runtime={'completed_utc':datetime.now(timezone.utc).isoformat(),'input_unchanged':True,'software':versions,'execution':'Linux container; image extraction, workbook production, openpyxl checks and business tests. LibreOffice rendering is recorded separately.','test_summary':summary[-1] if summary else 'No summary','test_exit_code':proc.returncode,'test_seconds':round(time.perf_counter()-test_start,3),'total_seconds':round(time.perf_counter()-start,3),'limitations':['Confidence is an OCR heuristic, not a probability of correctness.','Accuracy applies only to the supplied image set.']}
Path(a.output,'runtime.json').write_text(json.dumps(runtime,indent=2)+'\n')
if proc.returncode: raise SystemExit(proc.returncode)
print('PASS: scanned forms extraction and tests')
