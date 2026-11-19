"""Build source and cleaned-list extracts for Calc to render."""
import csv
import json
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.page import PageMargins
from workbook_meta import scrub_properties


def build():
    raw = list(openpyxl.load_workbook('input/input.xlsx')['Records'].values)
    with open('output/audit.csv') as f:
        audit = list(csv.DictReader(f))
    cleaned = list(openpyxl.load_workbook('output/customer-register.xlsx')['Records'].values)
    kept = [row for row in audit if row['status'] == 'kept'][:4]
    book = openpyxl.Workbook()
    book.remove(book.active)
    for title in ('Before', 'After'):
        ws = book.create_sheet(title)
        ws.sheet_view.showGridLines = False
        ws.merge_cells('A1:C1')
        ws['A1'] = 'BEFORE  /  mixed dates & amounts' if title == 'Before' else 'AFTER  /  cleaned customer list'
        ws.merge_cells('A2:C2')
        ws['A2'] = 'Opening entries · customer, date, amount'
        ws.append(['Customer', 'Date', 'Amount'])
        for index in range(4):
            record = raw[index+1] if title == 'Before' else cleaned[index+1]
            ws.append([record[2], record[3], record[4]])
            if title == 'After':
                ws.cell(index+4, 2).number_format = 'yyyy-mm-dd'
                ws.cell(index+4, 3).number_format = '#,##0.00;[Red](#,##0.00)'
        ws.merge_cells('A8:C8')
        ws['A8'] = 'Two matching entries · one impossible date' if title == 'Before' else 'Repeats removed · doubtful entries saved separately'
        for row in ws:
            for cell in row:
                cell.font = Font(name='Carlito', size=16, color='183E45')
                cell.alignment = Alignment(vertical='center', horizontal='left', indent=1)
                cell.fill = PatternFill('solid', fgColor='FFFFFF' if cell.row % 2 else 'EDF5F3')
                if cell.row in (1, 3):
                    cell.fill = PatternFill('solid', fgColor='183E45')
                    cell.font = Font(name='Carlito', size=16, color='FFFFFF', bold=True)
                if cell.row in (2, 8):
                    cell.font = Font(name='Carlito', size=13, color='42616A')
            ws.row_dimensions[row[0].row].height = 34 if row[0].row in (1, 8) else 31
        for column, width in [('A', 27), ('B', 18), ('C', 15)]:
            ws.column_dimensions[column].width = width
        ws.print_area = 'A1:C8'
        ws.page_setup.orientation = 'portrait'
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=.1, right=.1, top=.1, bottom=.1, header=0, footer=0)
    book.properties.creator = book.properties.lastModifiedBy = 'Genco Doganay'
    book.save('evidence/proof.xlsx')
    scrub_properties('evidence/proof.xlsx')
    # Separate print sources prevent a PDF exporter from including another sheet.
    for title in ('Before', 'After'):
        single = openpyxl.load_workbook('evidence/proof.xlsx')
        for sheet in list(single):
            if sheet.title != title:
                single.remove(sheet)
        path = Path('output') / ('proof-' + title.lower() + '.xlsx')
        single.save(path)
        scrub_properties(path)
    Path('evidence/proof-source.json').write_text(json.dumps({
        'source': 'input/input.xlsx', 'source_sheet': 'Records', 'source_rows': [2,3,4,5],
        'after': 'output/customer-register.xlsx', 'after_sheet': 'Records',
        'after_rows': [2,3,4,5], 'after_source_rows': [int(r['source_row']) for r in kept],
        'audit': 'output/audit.csv',
        'display': 'Opening source entries and opening cleaned Records. Removed repeats and entries needing review are absent from the cleaned list.',
        'matching_fields': ['customer_id', 'name', 'date', 'amount'],
        'name_comparison': {'first': raw[1][2], 'duplicate': raw[2][2],
                            'rule': 'Trim and collapse whitespace; compare names with casefold.'}
    }, indent=2) + '\n')


if __name__ == '__main__':
    build()
