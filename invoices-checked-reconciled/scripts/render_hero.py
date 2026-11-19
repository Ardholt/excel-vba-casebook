"""Render an uncropped invoice check from source rows and the cached ledger.

Run in the case container after generating the checked output. No amounts are
typed into the artwork: the source workbook, CSVs and ledger must agree.
"""
import csv
import hashlib
import json
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

from reconcile import reconcile, rounded


def records(name):
    with Path(f'output/{name}.csv').open(newline='') as stream:
        return list(csv.DictReader(stream))


def money(value):
    return f'£{Decimal(str(value)):.2f}'


def main():
    iid = 'INV-0017'
    result = reconcile(Path('input/input.xlsx'))
    invoice = next(r for r in result['invoices'] if r['invoice_id'] == iid)
    lines = [r for r in result['lines'] if r['invoice_id'] == iid]
    payments = [r for r in result['payments'] if r['invoice_id'] == iid]
    csv_invoice = next(r for r in records('invoices') if r['invoice_id'] == iid)
    for key in ('net', 'tax', 'gross', 'declared_total', 'difference', 'allocated', 'balance'):
        assert Decimal(csv_invoice[key]) == invoice[key], key
    for row in records('lines'):
        if row['invoice_id'] == iid:
            source = next(r for r in lines if r['line_id'] == row['line_id'])
            for key in ('net', 'tax', 'gross', 'quantity', 'unit_price', 'tax_rate'):
                assert Decimal(row[key]) == source[key], key
    assert len(lines) == 3 and len(payments) == 1
    payment = payments[0]
    assert payment['status'] == 'allocated'
    csv_payment = next(r for r in records('payments') if r['payment_id'] == payment['payment_id'])
    assert Decimal(csv_payment['amount']) == payment['amount'] == invoice['allocated']
    assert invoice['difference'] > 0 and invoice['balance'] == 0

    book = load_workbook('output/checked-ledger.xlsx', data_only=True)
    ledger = book['Ledger']
    ledger_row = next(r[0].row for r in ledger if r[0].value == iid)
    for col, key in ((2, 'net'), (3, 'tax'), (4, 'gross'), (5, 'declared_total'),
                     (6, 'difference'), (8, 'allocated'), (9, 'balance')):
        assert rounded(Decimal(str(ledger.cell(ledger_row, col).value))) == invoice[key], key
    assert ledger.cell(ledger_row, 7).value == 'REVIEW'
    book.close()

    # A narrow, complete proof card stays readable when GitHub scales to a phone.
    scale = 2
    canvas = Image.new('RGB', (680 * scale, 640 * scale), '#f5f7f6')
    draw = ImageDraw.Draw(canvas)
    ink, muted, green, amber = '#173e43', '#466168', '#e4efeb', '#fff0ce'
    fonts = Path('/usr/share/fonts/truetype/dejavu')

    def text(x, y, value, size=26, bold=False, color=ink, right=False):
        font = ImageFont.truetype(str(fonts / ('DejaVuSans-Bold.ttf' if bold else 'DejaVuSans.ttf')), size * scale)
        assert draw.textlength(value, font=font) <= (660 - x if not right else x - 20) * scale, value
        draw.text((x * scale, y * scale), value, font=font, fill=color, anchor='ra' if right else 'la')

    def rect(box, fill):
        draw.rectangle(tuple(int(v * scale) for v in box), fill=fill)

    rect((0, 0, 680, 66), ink)
    text(20, 15, f'{iid}  •  Invoice total needs review', 27, True, '#ffffff')
    rect((16, 82, 332, 384), '#ffffff')
    rect((348, 82, 664, 384), green)
    text(28, 94, 'SOURCE INVOICE', 25, True)
    text(360, 94, 'CHECKED LEDGER', 25, True)
    text(28, 130, 'input.xlsx extract', 22, color=muted)
    text(360, 130, f'Ledger row {ledger_row} • {iid}', 20, color=muted)
    text(28, 172, 'Qty × price', 24, True)
    text(318, 172, 'Tax', 24, True, right=True)
    for y, row in zip((211, 250, 289), lines):
        text(28, y, f'{row["quantity"]} × {money(row["unit_price"])}', 26)
        text(318, y, f'{row["tax_rate"] * 100:g}%', 26, right=True)
    text(28, 341, 'Stated', 25, True)
    text(318, 341, money(invoice['declared_total']), 26, True, right=True)
    for y, label, key in ((172, 'Net', 'net'), (211, 'Tax', 'tax'),
                          (250, 'Gross', 'gross'), (289, 'Payment', 'allocated'),
                          (341, 'Remaining', 'balance')):
        text(360, y, label, 25, bold=y in (250, 341))
        text(652, y, money(invoice[key]), 26, bold=y in (250, 341), right=True)

    text(20, 398, f'Net + tax: {money(invoice["net"])} + {money(invoice["tax"])} = {money(invoice["gross"])}', 26)
    rect((16, 444, 664, 514), amber)
    text(28, 451, 'Stated − calculated = difference to review', 23, True)
    text(28, 481, f'{money(invoice["declared_total"])} − {money(invoice["gross"])} = {money(invoice["difference"])}', 27, True)
    text(20, 530, f'Gross − payment ({payment["payment_id"]})', 24, True)
    text(20, 563, f'{money(invoice["gross"])} − {money(invoice["allocated"])} = {money(invoice["balance"])} remaining', 27, True)
    text(20, 606, 'Payment covered. Invoice difference still open.', 24)
    Path('media').mkdir(exist_ok=True)
    canvas.save('media/hero.png')

    source_files = ['input/input.xlsx', 'output/checked-ledger.xlsx',
                    'output/invoices.csv', 'output/lines.csv', 'output/payments.csv',
                    'output/exceptions.csv']
    proof = {
        'method': 'Typeset source extract beside a transposed ledger row; values verified against cached workbook and CSVs.',
        'invoice_id': iid,
        'source': {'file': 'input/input.xlsx', 'Invoices': invoice['source_row'],
                   'Lines': [r['source_row'] for r in lines], 'Payments': payment['source_row']},
        'ledger': {'file': 'output/checked-ledger.xlsx', 'sheet': 'Ledger', 'row': ledger_row},
        'invoice': invoice, 'lines': lines, 'payment': payment,
        'source_sha256': {p: hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in source_files},
        'render': {'file': 'media/hero.png', 'width': canvas.width, 'height': canvas.height},
    }
    Path('evidence/hero-proof.json').write_text(json.dumps(proof, default=str, indent=2) + '\n')
    media_path = Path('evidence/media.json')
    media = json.loads(media_path.read_text())
    media['method'] = 'LibreOffice sheet captures; hero typeset directly from source and checked ledger; PNG metadata stripped.'
    hero = next(r for r in media['images'] if r['file'] == '../media/hero.png')
    hero.update(sources=['../' + p for p in source_files], numbers='hero-proof.json',
                recipe='../scripts/render_hero.py')
    media_path.write_text(json.dumps(media, indent=2) + '\n')
    print('PASS: hero amounts match source invoice, line calculations, payment and cached ledger')


if __name__ == '__main__':
    main()
