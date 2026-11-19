"""Render the actual workbook print regions, retaining only clean PNG evidence."""
import hashlib
import json
import subprocess
import tempfile
from pathlib import Path
from PIL import Image
import openpyxl


def render(source, dest):
    dest = Path(dest)
    dest.parent.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.render-', dir=dest.parent) as tmp:
        subprocess.run(['xlrender', str(source), tmp, '--sheet', 'Dashboard'], check=True, capture_output=True, text=True)
        candidates = list(Path(tmp).glob('*.png'))
        if len(candidates) != 1:
            raise RuntimeError('Expected one dashboard render')
        with Image.open(candidates[0]) as original:
            clean = Image.new('RGB', original.size, 'white')
            clean.paste(original.convert('RGB'))
            clean.save(dest)
    return {'file': dest.as_posix(), 'source': str(source), 'width': clean.width,
            'height': clean.height, 'sha256': hashlib.sha256(dest.read_bytes()).hexdigest()}


if __name__ == '__main__':
    captures = [render('output/vba-after.xlsx', 'images/dashboard.png'),
                render('output/vba-before.xlsx', 'images/before.png')]
    with tempfile.TemporaryDirectory(prefix='.view-', dir='output') as tmp:
        path = Path(tmp) / 'breakdown.xlsx'
        book = openpyxl.load_workbook('output/vba-after.xlsx')
        report=json.loads(Path('output/report.json').read_text())
        end=13+max(1,len(report['paid_by_month']))+len(report['paid_by_category'])
        book['Dashboard'].print_area = f'A11:B{end}'
        book.save(path)
        capture = render(path, 'images/breakdown.png')
        capture['source'] = 'output/vba-after.xlsx'
        capture['region'] = f'Dashboard!A11:B{end}'
        captures.append(capture)
    Path('evidence/images.json').write_text(json.dumps({'renderer': 'LibreOffice through xlrender; macros disabled', 'captures': captures}, indent=2)+'\n')
