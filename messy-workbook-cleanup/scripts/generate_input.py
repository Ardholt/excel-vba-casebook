"""Rebuild the included inputs with a fixed, inspectable set of mistakes."""
import json
from datetime import date, timedelta
from pathlib import Path
import openpyxl
from normalize_records import FIELDS, table
from workbook_meta import scrub_properties


def generate(destination):
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    rows = [
        ['R001', ' 00071 ', '  ACCOUNT   N-01 ', '03/04/2026', '£1,240.50'],
        ['R002', '00071', 'account n-01', '2026-04-03', '1,240.50'],
        ['R003', '00082', 'Account W-01 ', '17/04/2026', '(£85.20)'],
        ['R004', '00093', 'Account S-01', '31/02/2026', '120.00'],
    ]
    for i in range(5, 24):
        day = date(2026, 4, 1) + timedelta(days=i)
        money = (i * 1379 + 213) / 100
        rows.append([f'R{i:03}', f'{100+i:05}', f'Account {i:03}',
                     day.isoformat() if i % 2 else day.strftime('%d/%m/%Y'),
                     f'{money:.2f}' if i % 3 else f'£{money:,.2f}'])
    for i, original in enumerate([rows[4], rows[7], rows[11], rows[16]], 24):
        duplicate = list(original)
        duplicate[0] = f'R{i:03}'
        duplicate[2] = '  ' + duplicate[2].upper() + '  '
        rows.append(duplicate)
    rows.extend([
        ['R028', '00228', 'Account 028', '2026-04-29', '12,34.50'],
        ['R029', '00229', 'Account 029', '2026-04-30', 'not entered'],
        ['R030', '00330', 'Account N-01', '03/04/2026', '£1,240.50'],
    ])
    # A quarter of customer activity, retaining the opening audit extract.
    added = []
    for i in range(450):
        day = date(2026, 4, 1) + timedelta(days=i % 90)
        cents = (i * 7919) % 89500 + 500
        if i % 17 == 0:
            cents = -cents
        magnitude = f'{abs(cents) // 100:,}.{abs(cents) % 100:02}'
        amount = f'(£{magnitude})' if cents < 0 else f'£{magnitude}'
        identifier = f'{500+i:05}'
        name = f'Account {500+i:04}'
        added.append([f'R{31+i:03}', ' ' + identifier + ' ' if i % 5 == 0 else identifier,
                      '  ' + name.upper() + '  ' if i % 4 == 0 else name,
                      day.isoformat() if i % 3 == 0 else day.strftime('%d/%m/%Y'), amount])
    rows.extend(added)
    for i in range(60):
        original = added[i * 7]
        duplicate = list(original)
        duplicate[0] = f'R{481+i:03}'
        duplicate[1] = duplicate[1].strip()
        duplicate[2] = duplicate[2].strip().lower()
        rows.append(duplicate)
    for i in range(10):
        rows.append([f'R{541+i:03}', f'{1000+i:05}', f'Account {1000+i:04}',
                     '31/06/2026' if i % 2 == 0 else '2026-06-30',
                     '40.00' if i % 2 == 0 else '12,34.50'])
    book = openpyxl.Workbook()
    book.remove(book.active)
    book.properties.creator = book.properties.lastModifiedBy = 'Genco Doganay'
    table(book, 'Records', FIELDS, rows)
    book.save(destination / 'input.xlsx')
    scrub_properties(destination / 'input.xlsx')
    (destination / 'rules.json').write_text(json.dumps({'date_order': 'DMY'}, indent=2) + '\n')


if __name__ == '__main__':
    generate('input')
