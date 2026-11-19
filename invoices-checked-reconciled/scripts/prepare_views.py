"""Export single-sheet copies for consistent sheet rendering."""
from pathlib import Path
from openpyxl import load_workbook
from run_pipeline import clean_properties

folder = Path('.render-tmp')
folder.mkdir(exist_ok=True)
for name, source, sheet in [('before','input/input.xlsx','Overview'), ('after','output/checked-ledger.xlsx','Payment list'), ('invoice','output/invoice-views.xlsx','Invoice 017')]:
    book = load_workbook(source)
    for tab in list(book.sheetnames):
        if tab != sheet:
            del book[tab]
    book.active = 0
    book[sheet].sheet_state = 'visible'
    if name == 'after':
        book[sheet].print_area = 'A1:C4'
        book[sheet].page_setup.fitToHeight = 1
    path = folder / f'{name}.xlsx'
    book.save(path)
    clean_properties(path)
    book.close()
