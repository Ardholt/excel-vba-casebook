from pathlib import Path
import tempfile
import zipfile
import json
from pyopenvba import ExcelFile
import xlsxwriter
from metadata import scrub

source = Path('vba/CombineFiles.bas').read_text()
with tempfile.TemporaryDirectory() as temp:
    raw = Path(temp) / 'package.xlsm'
    with ExcelFile.create_new(str(raw)) as book:
        book.set_module('Module1', source)
        book.save()
    with zipfile.ZipFile(raw) as z:
        binary = z.read('xl/vbaProject.bin')
    project = Path(temp) / 'vbaProject.bin'
    project.write_bytes(binary)
    wb = xlsxwriter.Workbook('combine.xlsm')
    wb.set_properties({'author': 'Genco Doganay', 'company': '', 'title': 'Refresh branch files', 'comments': ''})
    wb.add_vba_project(str(project))
    wb.set_vba_name('ThisWorkbook')
    cfg = wb.add_worksheet('Config')
    cfg.set_vba_name('Sheet1')
    heading = wb.add_format({'bold': True, 'bg_color': '#183F43', 'font_color': 'white', 'font_size': 16})
    cfg.write_row(0, 0, ['FOLDER REFRESH', 'Run RefreshFolder manually'], heading)
    cfg.write_row(1, 0, ['Source folder', 'input/files'])
    cfg.write_row(2, 0, ['Status', 'Ready'])
    cfg.write_row(3, 0, ['Rows combined', 0])
    cfg.write_row(5, 0, ['Input', 'Immediate folder; sheet named Data'])
    cfg.write_row(6, 0, ['Output', 'Combined and Manifest are replaced on refresh'])
    cfg.write_row(7, 0, ['Values', 'Source values and extra columns are retained'])
    cfg.write_row(8, 0, ['Saving', 'Save a separate copy after reviewing Manifest'])
    cfg.set_column('A:A', 24)
    cfg.set_column('B:B', 65)
    for i, name in enumerate(['Combined', 'Manifest'], 2):
        ws = wb.add_worksheet(name)
        ws.set_vba_name(f'Sheet{i}')
        ws.freeze_panes(1, 0)
        ws.set_column('A:J', 22)
    wb.close()
scrub('combine.xlsm')
print('Packaged readable VBA into combine.xlsm')
