"""Create the page figure from the source and result sheet renders."""
import json,subprocess
from pathlib import Path
from PIL import Image,ImageDraw
r=json.loads(Path('output/report.json').read_text())
source=json.loads(Path('evidence/hero-source.json').read_text())
Path('media').mkdir(exist_ok=True)
for name in ['Before','After']:
    subprocess.run(['convert',f'images/01-{name}.png','-background','white','-gravity','north','-resize','1280x406','-extent','1280x406','-strip',f'media/{name.lower()}.png'],check=True)
subprocess.run(['bash',str(Path(__file__).resolve().parents[2] / 'tools/hero/make-hero'),'media/before.png','media/after.png','media/hero.png','--title',f"{r['input_forms']} forms logged; unclear entries flagged",'--before-note','Faint item code on the scan','--after-note',f"Code left blank; {r['reviewed']['unresolved']} entries to check",'--before-where',source['image'],'--after-where','reviewed.xlsx','--tag','EXCEL & VBA · CASE 05'],check=True)
