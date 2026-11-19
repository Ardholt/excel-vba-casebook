"""Compose complete worksheet rows and their recorded changes for the README."""
import csv
import hashlib
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def main():
    with open('output/changes.csv') as source:
        changes = list(csv.DictReader(source))
    with open('output/duplicates.csv') as source:
        duplicates = list(csv.DictReader(source))
    edits = {r['field']: r for r in changes if r['source_row'] == '2'}
    repeat = next(r for r in duplicates if r['source_row'] == '3')
    assert ' '.join(edits['name']['before'].split()) == edits['name']['after']
    assert repeat['kept_source_row'] == '2'
    log = [
        f"Row 2 · spaces trimmed; date → {edits['date']['after']}",
        f"Row 2 · {edits['amount']['before']} → {edits['amount']['after']}",
        f"Row {repeat['source_row']} · repeat removed; row {repeat['kept_source_row']} kept",
    ]
    # These boundaries contain the header and two complete rows, inspected in
    # the 780px Calc renders. No worksheet values are redrawn or overlaid.
    crop = (0, 158, 780, 371)
    canvas = Image.new('RGB', (780, 724), '#ffffff')
    draw = ImageDraw.Draw(canvas)
    font_dir = '/usr/share/fonts/truetype/dejavu/'
    heading = ImageFont.truetype(font_dir + 'DejaVuSans-Bold.ttf', 29)
    body = ImageFont.truetype(font_dir + 'DejaVuSans.ttf', 28)
    for name, top, caption in [('before', 0, 'BEFORE · repeated entry'),
                               ('after', 272, 'AFTER · cleaned list')]:
        draw.rectangle((0, top, 779, top + 47), fill='#183e45')
        draw.text((16, top + 7), caption, font=heading, fill='white')
        with Image.open(f'images/{name}.png') as sheet:
            assert sheet.size == (780, 600), 'Inspect crop after worksheet layout changes'
            canvas.paste(sheet.crop(crop).convert('RGB'), (0, top + 48))
    draw.rectangle((0, 545, 779, 723), fill='#edf5f3')
    draw.text((16, 554), 'CHANGE LOG · source sheet rows', font=heading, fill='#183e45')
    for index, line in enumerate(log):
        assert draw.textbbox((0, 0), line, font=body)[2] <= 748
        draw.text((16, 596 + index * 37), line, font=body, fill='#183e45')
    canvas.save('media/hero.png')
    Path('evidence/hero.json').write_text(json.dumps({
        'file': 'media/hero.png', 'renderer': 'scripts/build_hero.py; Pillow',
        'source_images': ['images/before.png', 'images/after.png'],
        'source_sha256': {p: digest(p) for p in [
            'images/before.png', 'images/after.png', 'output/changes.csv',
            'output/duplicates.csv', 'output/report.json']},
        'crop_each': {'x': 0, 'y': 158, 'width': 780, 'height': 213},
        'canvas_pixels': list(canvas.size),
        'before_source_sheet_rows': [2, 3], 'after_cleaned_sheet_rows': [2, 3],
        'after_original_sheet_rows': [2, 4], 'trace': 'proof-source.json',
        'change_log': log, 'typography_licences': 'media/ATTRIBUTION.md',
        'sha256': digest('media/hero.png'),
    }, indent=2) + '\n')


if __name__ == '__main__':
    main()
