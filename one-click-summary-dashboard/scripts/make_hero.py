"""Render the current summary with recent months and every category visible."""
import hashlib
import json
import tempfile
from pathlib import Path
import openpyxl
from render_evidence import render

source=Path('output/vba-after.xlsx')
report=json.loads(Path('output/report.json').read_text())
months=list(report['paid_by_month'])
with tempfile.TemporaryDirectory(prefix='.summary-',dir='output') as tmp:
    book=openpyxl.load_workbook(source)
    sheet=book['Dashboard']
    # Presentation-only view: retain the actual saved macro values and formatting.
    for row in range(12,12+max(0,len(months)-4)):
        sheet.row_dimensions[row].hidden=True
    sheet['A11']='Recent months'
    end=13+len(months)+len(report['paid_by_category'])
    sheet.print_area=f'A11:B{end}'
    path=Path(tmp)/'summary.xlsx'
    book.save(path)
    capture=render(path,'media/hero.png')
capture.update(source=source.as_posix(),source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
               region=f'Dashboard!A11:B{end}',visible_months=months[-4:],
               categories=list(report['paid_by_category']),
               presentation='Earlier months hidden; month heading shortened. All figures are unchanged saved VBA results.',
               figures='output/report.json')
Path('evidence/hero.json').write_text(json.dumps(capture,indent=2)+'\n')
