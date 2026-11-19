"""Render a compact, file-backed approval diagram with Chromium.

Run from the case folder after scripts/inner.sh. This is a typeset extract,
not a screenshot of an interactive approval control.
"""
import csv
import subprocess
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
from customers import read_customers


def csv_rows(path):
    with Path(path).open(newline='', encoding='utf-8') as stream:
        return list(csv.DictReader(stream))


source, target = 'C0403', 'C0003'
records = {r['customer_id']: r for r in read_customers('input/input.xlsx')}
approval = next(r for r in csv_rows('input/confirmed.csv')
                if r['source_id'] == source and r['target_id'] == target)
audit = next(r for r in csv_rows('output/audit.csv') if r['source_id'] == source)
assert audit['retained_id'] == target
assert audit['confirmation_path'] == f'{source} > {target}'
assert audit['reason'] == 'Explicit confirmation'
book = openpyxl.load_workbook('output/result.xlsx', read_only=True, data_only=True)
kept = {row[0]: row for row in list(book['Customers'].values)[1:]}
assert source not in kept and kept[target][1] == records[target]['name']
assert (source, target) in list(book['Confirmed'].values)[1:]
book.close()

parts = ['<svg xmlns="http://www.w3.org/2000/svg" width="720" height="640" viewBox="0 0 720 640">',
         '<rect width="720" height="640" fill="#10100f"/>']


def text(x, y, value, size=27, color='#f4f1eb', mono=False, bold=False):
    family = 'Liberation Mono' if mono else 'Liberation Sans'
    weight = 'bold' if bold else 'normal'
    parts.append(f'<text x="{x}" y="{y}" font-family="{family}" font-size="{size}" '
                 f'font-weight="{weight}" fill="{color}">{escape(value)}</text>')


def line(y):
    parts.append(f'<path d="M28 {y} H692" stroke="#575750"/>')


text(28, 43, 'TWO NAMES', 25, bold=True)
text(28, 77, 'input/input.xlsx', 23, '#c4c1b8', mono=True)
for y, key in [(122, target), (162, source)]:
    text(28, y, key, 26, mono=True)
    text(140, y, records[key]['name'], 27)
line(187)

text(28, 228, 'SAVED APPROVAL', 25, '#b8e7ce', bold=True)
text(28, 262, 'input/confirmed.csv', 23, '#c4c1b8', mono=True)
parts.append('<rect x="28" y="280" width="664" height="135" rx="5" fill="#eff6f5"/>')
text(46, 313, 'Old number', 24, '#183d4a')
text(370, 313, 'Number to keep', 24, '#183d4a')
text(46, 350, 'source_id', 25, '#183d4a', mono=True)
text(370, 350, 'target_id', 25, '#183d4a', mono=True)
text(46, 394, approval['source_id'], 34, '#183d4a', mono=True, bold=True)
text(303, 392, '→', 34, '#183d4a')
text(370, 394, approval['target_id'], 34, '#183d4a', mono=True, bold=True)

text(28, 461, 'KEPT IN THE RESULT', 25, bold=True)
text(28, 495, 'output/result.xlsx', 23, '#c4c1b8', mono=True)
text(28, 538, target, 26, mono=True)
text(140, 538, kept[target][1], 27)
line(562)
text(28, 595, 'output/audit.csv', 23, '#c4c1b8', mono=True)
text(28, 626, f'{source} → {target} · {audit["reason"]}', 25)
parts.append('</svg>')
Path('media/hero.svg').write_text('\n'.join(parts) + '\n')
subprocess.run(['/opt/page/bin/python', '-c', '''
from pathlib import Path
from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=True, args=['--no-sandbox'])
    page = browser.new_page(viewport={'width': 720, 'height': 640}, device_scale_factor=2)
    page.goto(Path('media/hero.svg').resolve().as_uri())
    page.evaluate('document.fonts.ready')
    page.screenshot(path='media/hero.png')
    browser.close()
'''], check=True)
print('PASS: hero approval and result match the input, confirmation file and audit')
