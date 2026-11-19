"""Render a compact, data-backed proof summary, then capture the README."""
import csv
import hashlib
import json
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook


def csv_rows(path):
    with Path(path).open(newline='') as handle:
        return list(csv.DictReader(handle))


def money(value):
    return f'{Decimal(value):,.2f}'


a = next(r for r in csv_rows('input/layout-a.csv') if r['id'] == 'A002')
c = next(r for r in csv_rows('input/layout-c.csv') if r['reference'] == 'C002')
book = load_workbook('input/layout-b.xlsx', read_only=True, data_only=True)
values = iter(book['Transactions'].values)
headers = next(values)
b = next(r for r in (dict(zip(headers, row)) for row in values) if r['id'] == 'B002')
book.close()
ledger = csv_rows('output/ledger.csv')
selected = [next(r for r in ledger if r['id'] == key) for key in ('A002', 'B002', 'C002')]
checks = csv_rows('output/reconciliation.csv')
report = json.loads(Path('output/report.json').read_text())
assert len(ledger) == report['output_records']
assert [Decimal(r['amount']) for r in selected] == [
    Decimal(a['credit'] or '0') - Decimal(a['debit'] or '0'),
    Decimal(b['amount']),
    Decimal(c['amount']) * (-1 if c['direction'] == 'OUT' else 1),
]

image = Image.new('RGB', (720, 890), '#ffffff')
draw = ImageDraw.Draw(image)
ink, muted, rule = '#183541', '#4d616b', '#d6e0e4'
font_dir = Path('/usr/share/fonts/truetype/dejavu')


def text(x, y, value, size=30, bold=False, color=ink, right=False):
    font = ImageFont.truetype(str(font_dir / ('DejaVuSans-Bold.ttf' if bold else 'DejaVuSans.ttf')), size)
    width = draw.textlength(value, font=font)
    if right:
        x -= width
    assert 0 <= x and x + width <= 700, (value, x, width)
    draw.text((x, y), value, font=font, fill=color)


draw.rectangle((0, 0, 719, 889), outline=rule, width=2)
text(22, 16, 'THREE SOURCE FORMATS', 32, True)
text(22, 58, 'Selected money columns · amounts in £', 27, color=muted)
for y, title, detail in [
    (104, 'A · layout-a.csv', f"{a['id']}   debit {a['debit']}   credit {a['credit'] or '—'}"),
    (187, 'B · layout-b.xlsx', f"{b['id']}   amount {b['amount']}"),
    (270, 'C · layout-c.csv', f"{c['reference']}   amount {c['amount']}   {c['direction']}"),
]:
    text(22, y, title, 29, True)
    text(22, y + 36, detail, 29)

draw.rectangle((1, 355, 718, 561), fill='#edf5f5')
text(22, 367, f"ONE LEDGER · {report['output_records']} rows kept", 32, True)
text(22, 414, 'Row ID', 28, color=muted)
text(220, 414, 'Account', 28, color=muted)
text(696, 414, 'Amount £', 28, color=muted, right=True)
for y, row in zip((452, 487, 522), selected):
    text(22, y, row['id'], 29)
    text(220, y, row['account'], 29)
    text(696, y, money(row['amount']).replace('-', '−'), 29, right=True)

text(22, 581, 'ALL ACCOUNT CLOSES · £', 32, True)
text(22, 627, 'Account', 28, color=muted)
text(465, 627, 'Statement', 28, color=muted, right=True)
text(696, 627, 'Ledger', 28, color=muted, right=True)
for y, row in zip((669, 718, 767), checks):
    review = row['status'] != 'OK'
    if review:
        draw.rectangle((1, y - 4, 718, y + 42), fill='#fff0d9')
    account = row['account'].removeprefix('ACCT-')
    text(22, y, f"{account} · {row['status']}", 28, True)
    text(465, y, money(row['declared_close']), 30, right=True)
    text(696, y, money(row['calculated_close']), 30, right=True)
delta = Decimal(checks[2]['closing_difference'])
assert delta < 0 and checks[2]['status'] == 'REVIEW'
text(22, 833, f'C is £{abs(delta):.2f} short · review required', 31, True, '#8a4600')
image.save('media/hero.png', optimize=True)

sources = ['input/layout-a.csv', 'input/layout-b.xlsx', 'input/layout-c.csv',
           'output/ledger.csv', 'output/reconciliation.csv', 'output/report.json']
Path('evidence/hero.json').write_text(json.dumps({
    'method': 'Pillow typeset summary of selected source columns, matching ledger rows and all account closes; not a worksheet screenshot.',
    'dimensions': list(image.size),
    'source_sha256': {p: hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in sources},
    'rendered_sha256': {'media/hero.png': hashlib.sha256(Path('media/hero.png').read_bytes()).hexdigest()},
    'script_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
    'ledger_rows': selected,
    'account_checks': checks,
    'source_selection': 'Row 3 including the header in each source; A002, B002, C002. Blank credit shown as an em dash. Account A/B/C abbreviates ACCT-A/B/C.',
    'fonts': 'DejaVu Sans, regular and bold; container system fonts',
}, indent=2) + '\n')
subprocess.run([sys.executable, 'scripts/capture_page.py'], check=True)
