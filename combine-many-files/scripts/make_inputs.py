"""Rebuild the supplied branch exports deterministically."""
from pathlib import Path
from datetime import date
import json
import sys
sys.path.insert(0, str(Path(__file__).parent))
from combine import clean_properties, style
from metadata import scrub
import openpyxl

folder = Path('input/files')
folder.mkdir(parents=True, exist_ok=True)
for i in range(1, 51):
    branch = (i-1) % 6 + 1
    month = (i-1) // 6 + 1
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Data' if i != 50 else 'Notes'
    headers = ['record_id', 'date', 'category', 'amount', 'branch']
    if i % 3 == 0:
        headers += ['channel']
    if i % 5 == 0:
        headers += ['batch_code']
    if i % 2 == 0:
        headers = headers[3:] + headers[:3]
    if i == 49:
        headers.remove('amount')
    if i == 50:
        ws.append(['Export not ready'])
    else:
        ws.append(headers)
        if i <= 46:
            for j in range(1, 121):
                # Credits and two distinct identical-looking records remain separate.
                amount = round((800 + (i * 71 + j * 37) % 28000) / 100, 2)
                if j % 29 == 0:
                    amount = -amount
                row = {'record_id': f'B{branch:02d}-{month:02d}-{j:04d}',
                       'date': date(2025, month, (j-1) % 28+1).isoformat(),
                       'category': ['Parts', 'Service', 'Delivery'][j % 3], 'amount': amount,
                       'branch': f'B{branch:02d}', 'channel': ['Counter', 'Online'][j % 2],
                       'batch_code': f'RUN-{month:02d}'}
                if j in (119, 120):
                    row.update(date=date(2025, month, 28).isoformat(), category='Parts', amount=42.50)
                ws.append([row[h] for h in headers])
        elif i == 49:
            ws.append(['B01-09-0001' if h == 'record_id' else '2025-09-01' if h == 'date' else 'Parts' if h == 'category' else 'B01' for h in headers])
    style(ws)
    clean_properties(wb)
    path = folder / f'{i:02d}_2025-{month:02d}_B{branch:02d}.xlsx'
    wb.save(path)
    scrub(path)
(folder / '~$01_2025-01_B01.xlsx').write_bytes(b'Workbook locked for editing')
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Before'
for row in [['SEPARATE BRANCH EXPORTS', ''], ['Files to bring together', 50],
            ['Column order', 'Changes between files'], ['Extra fields', 'Branch + other fields'],
            ['Empty or invalid files', 'Need a visible decision'], ['Source trail', 'Keep filename + row']]:
    ws.append(row)
from openpyxl.styles import Font, Alignment
for row in ws:
    for c in row:
        c.font = Font(name='Calibri', size=18, bold=c.row in (1,2), color='743E28')
        c.alignment = Alignment(vertical='center')
    ws.row_dimensions[row[0].row].height = 36
ws.merge_cells('A1:B1')
ws.column_dimensions['A'].width = 38
ws.column_dimensions['B'].width = 30
ws.sheet_view.showGridLines = False
ws.print_area = 'A1:B6'
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = 'landscape'
ws.page_setup.paperSize = ws.PAPERSIZE_A5
ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
clean_properties(wb)
wb.save('input/before.xlsx')
scrub('input/before.xlsx')
