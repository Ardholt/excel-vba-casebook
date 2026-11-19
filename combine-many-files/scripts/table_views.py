"""Build readable views directly from source headers and finished table rows."""
from pathlib import Path
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from combine import clean_properties
from metadata import scrub

out = Path('output/table-views.xlsx')
wb = openpyxl.Workbook()
before = wb.active
before.title = 'Before'
before.append(['File', 'Same fields, different order', None, None])
before.merge_cells('B1:D1')
names = ['01_2025-01_B01.xlsx', '02_2025-01_B02.xlsx']
labels = {'date': 'Date', 'amount': 'Amount', 'branch': 'Shop'}
source_evidence = []
for name in names:
    book = openpyxl.load_workbook(Path('input/files') / name, read_only=True, data_only=True)
    headers = list(next(book['Data'].values))
    chosen = [h for h in headers if h in labels]
    before.append([name] + [labels[h] for h in chosen])
    source_evidence.append({'file': name, 'original_headers': headers, 'visible_headers': chosen})
    book.close()

after = wb.create_sheet('Combined')
after.append(['Date', 'Amount', 'From file'])
book = openpyxl.load_workbook('output/consolidated.xlsx', read_only=True, data_only=True)
iterator = book['Combined'].iter_rows(values_only=True)
headers = list(next(iterator))
chosen_rows = []
for workbook_row, values in enumerate(iterator, 2):
    row = dict(zip(headers, values))
    if row['source_file'] in names and row['source_row'] == 2:
        source = openpyxl.load_workbook(Path('input/files') / row['source_file'], read_only=True, data_only=True)
        original = list(source['Data'].values)
        record = dict(zip(original[0], original[row['source_row'] - 1]))
        assert all(row[k] == record[k] for k in ('date','amount','record_id'))
        after.append([row['date'], row['amount'], row['source_file']])
        chosen_rows.append({'combined_sheet_row': workbook_row, 'source_file': row['source_file'],
                            'source_row': row['source_row'], 'date': row['date'], 'amount': row['amount'],
                            'record_id': row['record_id']})
        source.close()
book.close()
assert len(chosen_rows) == 2

for sheet, widths, size in [(before, [34, 13, 12, 11], 16), (after, [17, 16, 36], 18)]:
    for i, width in enumerate(widths, 1):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
    for row in sheet:
        for cell in row:
            cell.font = Font(name='Calibri', size=size, color='183F43', bold=cell.row == 1)
            cell.alignment = Alignment(vertical='center', horizontal='left')
            cell.border = Border(bottom=Side(style='thin', color='D7E5E2'))
            if cell.row == 1:
                cell.fill = PatternFill('solid', fgColor='E4EEEB')
        sheet.row_dimensions[row[0].row].height = 32
    sheet.sheet_view.showGridLines = False
    sheet.print_area = sheet.dimensions
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = 'landscape'
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = sheet.page_setup.fitToHeight = 1
for cell in after['B'][1:]:
    cell.number_format = '#,##0.00'
    cell.alignment = Alignment(vertical='center', horizontal='right', indent=1)
for cell in after['C']:
    cell.alignment = Alignment(vertical='center', horizontal='left', indent=1)
clean_properties(wb)
wb.save(out)
scrub(out)
import shutil
shutil.rmtree('.render-tmp', ignore_errors=True)
Path('.render-tmp').mkdir(exist_ok=True)
for sheet, name in [('Before','before'),('Combined','after')]:
    single = openpyxl.load_workbook(out)
    for other in list(single.sheetnames):
        if other != sheet:
            del single[other]
    target = Path('.render-tmp') / (name + '.xlsx')
    single.save(target)
    scrub(target)
Path('evidence/table-preview.json').write_text(json.dumps({
    'source': 'output/consolidated.xlsx', 'sheet': 'Combined',
    'view': 'First entry from each of the first two shop files; original values retained.',
    'header_labels': {'Date':'date','Amount':'amount','From file':'source_file'},
    'rows': chosen_rows, 'before': source_evidence,
    'checks': {'visible_values_match_source_and_combined': True}
}, indent=2)+'\n')
