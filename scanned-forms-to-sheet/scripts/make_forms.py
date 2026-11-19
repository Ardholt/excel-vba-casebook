"""Create the reproducible form set; never imported by the extractor."""
from pathlib import Path
import csv, json, zipfile
from PIL import Image, ImageDraw, ImageFont, ImageFilter

FONT = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
FIELDS = ['Form ID', 'Item code', 'Quantity', 'Amount']

def draw_form(path, values, damaged=(), hard=None):
    hard = hard or {}
    im = Image.new('RGB', (1000, 600), '#faf9f5')
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, 1000, 94), fill='#173d42')
    d.text((42, 25), 'ITEM ENTRY', font=ImageFont.truetype(FONT, 38), fill='white')
    for i, (label, value) in enumerate(zip(FIELDS, values)):
        y = 142 + 100*i
        d.text((45, y+10), label, font=ImageFont.truetype(FONT, 29), fill='#173d42')
        d.rectangle((400, y, 955, y+78), outline='#acb8b4', width=2)
        d.text((425, y+10), value, font=ImageFont.truetype(FONT, 40), fill='#b8b6af' if hard.get(label)=='faint' else '#151c1c')
        if hard.get(label)=='blurred':
            region=(410,y+5,945,y+72)
            im.paste(im.crop(region).filter(ImageFilter.GaussianBlur(1.3)),region)
            d=ImageDraw.Draw(im)
        if label in damaged:
            # Overwriting and ink loss obscure the actual printed characters.
            d.rectangle((420, y+5, 735, y+72), fill='#f5f3eb')
            for offset in range(0, 50, 7):
                d.line((430, y+12+offset, 710, y+65-offset), fill='#444843', width=4)
    im.save(path)

def main():
    dest = Path('input/forms'); dest.mkdir(parents=True, exist_ok=True)
    Path('validation').mkdir(exist_ok=True)
    rows, reviews, manifest = [], [], []
    for n in range(1, 49):
        values = [f'F{240+n}', f'IT-{620+n}', str(13+n), f'{87+n*19}.{(n*13)%100:02d}']
        position = (n-1)%12+1
        damaged = {4: ['Amount'], 8: ['Item code'], 11: ['Quantity']}.get(position, [])
        hard = {6: {'Item code':'faint'}, 9: {'Quantity':'blurred'}}.get(position,{})
        name = f'entry-{n:02}.png'
        draw_form(dest/name, values, damaged, hard)
        rows.extend({'image': name, 'field': f, 'value': v} for f, v in zip(FIELDS, values))
        manifest.append({'image':name,'condition':'obscured' if damaged else 'hard to read' if hard else 'clear','affected_fields':damaged or list(hard)})
        for field in damaged:
            if field!='Quantity': reviews.append([values[0],field,values[FIELDS.index(field)]])
    with open('validation/answer-key.csv', 'w', newline='') as f:
        w=csv.DictWriter(f,fieldnames=['image','field','value']);w.writeheader();w.writerows(rows)
    with open('input/review.csv','w',newline='') as f:
        w=csv.writer(f);w.writerow(['form_id','field','value']);w.writerows(reviews)
    Path('validation/batch.json').write_text(json.dumps({'forms':len(manifest),'fields':len(rows),'conditions':{k:sum(r['condition']==k for r in manifest) for k in ['clear','hard to read','obscured']},'supplied_confirmations':len(reviews),'files':manifest},indent=2)+'\n')
    with zipfile.ZipFile('input/forms.zip','w',zipfile.ZIP_DEFLATED) as z:
        for row in manifest: z.write(dest/row['image'],row['image'])
if __name__ == '__main__': main()
