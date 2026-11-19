"""Build a phone-readable comparison from scan pixels and the complete saved row."""
import hashlib
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook

source = json.loads(Path('evidence/hero-source.json').read_text())
book = load_workbook('output/reviewed.xlsx', data_only=True)
row = next(row for row in book['Data'].values if row[0] == source['image'])
values = [str(row[1]), row[2] or '', str(row[3]), f'{row[4]:.2f}']
assert values == [field['value'] for field in source['fields']]

font_path = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
bold_path = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
font = ImageFont.truetype(font_path, 30)
value_font = ImageFont.truetype(font_path, 36)
heading = ImageFont.truetype(bold_path, 30)
ink, teal, line = '#173d42', '#173d42', '#d5dedc'
canvas = Image.new('RGB', (720, 410), 'white')
draw = ImageDraw.Draw(canvas)
draw.rectangle((0, 0, 719, 67), fill=teal)
for x, label in [(16, 'Field'), (210, 'Scan'), (476, 'Sheet')]:
    draw.text((x, 16), label, font=heading, fill='white')

with Image.open(Path('input/forms') / source['image']) as scan:
    for index, (field, value) in enumerate(zip(source['fields'], values)):
        top = 68 + index * 76
        draw.line((0, top, 720, top), fill=line, width=2)
        draw.text((16, top + 21), ['Form no.', 'Item code', 'Quantity', 'Amount'][index], font=font, fill=ink)
        crop = scan.crop(tuple(map(int, field['rectangle'].split(','))))
        # Trim only the empty right margin, retaining the full ink and its faintness.
        crop = crop.crop((0, 0, 330, crop.height))
        crop.thumbnail((246, 60), Image.Resampling.LANCZOS)
        canvas.paste(crop, (210, top + (76 - crop.height) // 2))
        if not value:
            draw.rectangle((466, top + 1, 719, top + 75), fill='#f7e4b4')
        draw.text((476, top + 17), value or 'Blank*', font=value_font, fill=ink)

draw.text((16, 378), '* Check item code against the scan.', font=ImageFont.truetype(font_path, 27), fill=ink)
Path('media').mkdir(exist_ok=True)
canvas.save('media/hero.png')
source.update({
    'presentation': 'Four scan crops beside all four saved fields, arranged vertically for narrow screens. Blank* labels an empty cell; it is not a saved value.',
    'value_source': 'output/reviewed.xlsx',
    'worksheet': 'Data',
    'workbook_row': next(index for index, candidate in enumerate(book['Data'].values, 1) if candidate[0] == source['image']),
    'source_sha256': {str(path): hashlib.sha256(path.read_bytes()).hexdigest() for path in [Path('input/forms') / source['image'], Path('output/reviewed.xlsx')]},
})
Path('evidence/hero-source.json').write_text(json.dumps(source, indent=2) + '\n')
