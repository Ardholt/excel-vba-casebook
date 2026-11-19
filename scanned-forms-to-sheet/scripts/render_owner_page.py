"""Rebuild the worksheet evidence views and record the rendered sources."""
import hashlib,json,subprocess,time,shutil
from pathlib import Path
from PIL import Image

def run(*args):
    start=time.perf_counter()
    subprocess.run(args,check=True)
    return round(time.perf_counter()-start,3)

run('python3','scripts/build_hero_sheets.py')
evidence=json.loads(Path('evidence/rendering.json').read_text())
evidence['sheets']=[x for x in evidence['sheets'] if x['image'] not in ('images/01-Before.png','images/01-After.png')]
for name in ['Before','After']:
    dest=Path('render-temp')/name
    seconds=run('xlrender',f'output/hero-{name.lower()}.xlsx',str(dest))
    src=dest/f'01-{name}.png';out=Path('images')/src.name
    shutil.copyfile(src,out)
    with Image.open(out) as im:pixels=list(im.size)
    evidence['sheets'].append({'image':str(out),'workbook':f'output/hero-{name.lower()}.xlsx','pixels':pixels,'render_seconds':seconds,'sha256':hashlib.sha256(out.read_bytes()).hexdigest()})
run('python3','scripts/make_page_assets.py')
with Image.open('media/hero.png') as im:evidence['hero']['pixels']=list(im.size)
evidence['hero']['sha256']=hashlib.sha256(Path('media/hero.png').read_bytes()).hexdigest()
evidence['method']='LibreOffice sheet export and Poppler rasterisation for sheet views; Pillow composition of actual scan crops and openpyxl-read workbook values for the compact page figure.'
evidence['hero']['source_views']=['input/forms/'+json.loads(Path('evidence/hero-source.json').read_text())['image'],'output/reviewed.xlsx']
evidence['hero']['template']='Four field comparison rows, sized for a narrow GitHub README; scripts/make_page_assets.py.'
Path('evidence/rendering.json').write_text(json.dumps(evidence,indent=2)+'\n')
shutil.rmtree('render-temp')
run('python3','scripts/write_page.py')
run('python3','scripts/preview_page.py')
run('python3','scripts/record_final_tests.py')
runtime=json.loads(Path('evidence/runtime.json').read_text())
runtime['page_preview']=json.loads(Path('evidence/page-preview.json').read_text())
runtime['post_render_tests']=json.loads(Path('evidence/final-tests.json').read_text())
Path('evidence/runtime.json').write_text(json.dumps(runtime,indent=2)+'\n')
