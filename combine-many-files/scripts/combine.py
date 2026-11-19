"""Consolidate the immediate folder; keep each source record and its origin."""
import argparse
import csv
import hashlib
import json
import time
from pathlib import Path
from decimal import Decimal
import openpyxl
from metadata import scrub
from openpyxl.styles import Font, PatternFill, Alignment

REQUIRED = ['record_id', 'date', 'category', 'amount']
ORIGIN = ['source_file', 'source_row']
RESERVED = set(ORIGIN)
OUTPUT_NAMES = {'consolidated.xlsx', 'vba-result.xlsx', 'combine.xlsm'}


def style(ws, widths=None):
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    for c in ws[1]:
        c.fill = PatternFill('solid', fgColor='183F43')
        c.font = Font(name='Calibri', size=12, color='FFFFFF', bold=True)
        c.alignment = Alignment(wrap_text=True)
    ws.row_dimensions[1].height = 30
    for col in ws.columns:
        name = col[0].value
        ws.column_dimensions[col[0].column_letter].width = (widths or {}).get(name, 20)
    ws.sheet_view.showGridLines = False


def clean_properties(wb):
    wb.properties.creator = 'Genco Doganay'
    wb.properties.lastModifiedBy = 'Genco Doganay'
    wb.properties.title = 'Combine branch files'
    wb.properties.description = ''


def collect(folder):
    rows, manifest, extras = [], [], []
    for path in sorted(Path(folder).iterdir(), key=lambda p: p.name):
        if not path.is_file():
            continue
        if path.name.startswith('~$') or path.name.lower() in OUTPUT_NAMES:
            manifest.append([path.name, 'ignored', 0, 'Temporary or output file'])
            continue
        if path.suffix.lower() != '.xlsx':
            manifest.append([path.name, 'ignored', 0, 'Not an .xlsx workbook'])
            continue
        wb = None
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            if 'Data' not in wb.sheetnames:
                raise ValueError('Missing Data sheet')
            ws = wb['Data']
            iterator = ws.iter_rows(values_only=True)
            headers = [str(v).strip() if v is not None else '' for v in next(iterator, ())]
            while headers and headers[-1] == '':
                headers.pop()
            if not headers or any(not h for h in headers):
                raise ValueError('Blank header')
            if len(set(headers)) != len(headers):
                raise ValueError('Duplicate header')
            if any(h in RESERVED for h in headers):
                raise ValueError('Reserved provenance header')
            if any(h not in headers for h in REQUIRED):
                raise ValueError('Missing required columns')
            for h in headers:
                if h not in REQUIRED and h not in extras:
                    extras.append(h)
            count = 0
            for source_row, cells in enumerate(iterator, 2):
                if all(v is None or v == '' for v in cells):
                    continue
                if any(v is not None and v != '' for v in cells[len(headers):]):
                    raise ValueError('Value without header')
                row = dict(zip(headers, cells))
                row.update(source_file=path.name, source_row=source_row)
                rows.append(row)
                count += 1
            manifest.append([path.name, 'accepted' if count else 'empty', count, ''])
        except Exception as exc:
            rows = [r for r in rows if r['source_file'] != path.name]
            reason = str(exc) if isinstance(exc, ValueError) else 'Unreadable workbook'
            manifest.append([path.name, 'rejected', 0, reason])
        finally:
            if wb:
                wb.close()
    return REQUIRED + ORIGIN + extras, rows, manifest


def combine(folder, output):
    started = time.perf_counter()
    folder, output = Path(folder).resolve(), Path(output).resolve()
    if not folder.is_dir():
        raise ValueError('Input folder does not exist')
    if folder == output or folder in output.parents or output in folder.parents:
        raise ValueError('Input and output must be separate, non-nested folders')
    headers, rows, manifest = collect(folder)
    output.mkdir(parents=True, exist_ok=True)
    statuses = {s: sum(r[1] == s for r in manifest) for s in ['accepted', 'empty', 'rejected', 'ignored']}
    total = sum((Decimal(str(r['amount'])) for r in rows), Decimal(0))
    business = {'source_workbooks': sum(r[1] != 'ignored' for r in manifest), 'files': statuses,
                'output_rows': len(rows), 'amount_total': str(total.quantize(Decimal('.01'))),
                'columns': headers, 'duplicate_policy': 'Preserve source records; no deduplication'}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Combined'
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h) for h in headers])
    for cell in ws['D'][1:]:
        cell.number_format = '#,##0.00;[Red](#,##0.00)'
    style(ws, {'source_file': 32})
    for name, content in [('Manifest', manifest), ('Rejected', [r for r in manifest if r[1] == 'rejected'])]:
        sheet = wb.create_sheet(name)
        sheet.append(['file', 'status', 'rows', 'reason'])
        for row in content:
            sheet.append(row)
        style(sheet, {'file': 35, 'reason': 34})
    proof = wb.create_sheet('Refresh check')
    for row in [['ONE TABLE · EVERY ROW TRACEABLE', ''], ['Source workbooks', business['source_workbooks']],
                ['Rows combined', len(rows)], ['Files needing attention', statuses['rejected']],
                ['Amount total', float(total)], ['Refresh', 'Replace the table']]:
        proof.append(row)
    proof.merge_cells('A1:B1')
    proof.column_dimensions['A'].width = 38
    proof.column_dimensions['B'].width = 25
    for row in proof:
        for c in row:
            c.font = Font(name='Calibri', size=18, bold=c.row in (1, 3), color='183F43')
            c.alignment = Alignment(vertical='center')
        proof.row_dimensions[row[0].row].height = 36
    proof['B5'].number_format = '#,##0.00'
    proof['B3'].number_format = '#,##0'
    proof.sheet_view.showGridLines = False
    proof.print_area = 'A1:B6'
    proof.sheet_properties.pageSetUpPr.fitToPage = True
    proof.page_setup.orientation = 'landscape'
    proof.page_setup.paperSize = proof.PAPERSIZE_A5
    proof.page_setup.fitToWidth = 1
    proof.page_setup.fitToHeight = 1
    clean_properties(wb)
    wb.save(output / 'consolidated.xlsx')
    scrub(output / 'consolidated.xlsx')
    for filename, content in [('manifest.csv', manifest), ('rejections.csv', [r for r in manifest if r[1] == 'rejected'])]:
        with (output / filename).open('w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['file', 'status', 'rows', 'reason'])
            writer.writerows(content)
    with (output / 'combined.csv').open('w', newline='') as f:
        writer = csv.DictWriter(f, headers)
        writer.writeheader()
        writer.writerows(rows)
    hashes = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(folder.iterdir()) if p.is_file()}
    report = {**business, 'source_sha256': hashes, 'reference_seconds': round(time.perf_counter() - started, 3),
              'checks': {'manifest_rows_equal_output': sum(r[2] for r in manifest) == len(rows),
                         'unique_provenance': len({(r['source_file'], r['source_row']) for r in rows}) == len(rows)},
              'scope': 'Immediate folder; values only; exact case-sensitive headers trimmed at edges'}
    (output / 'report.json').write_text(json.dumps(report, indent=2) + '\n')
    (output / 'evidence.txt').write_text(
        f"Source workbooks: {business['source_workbooks']}\nAccepted: {statuses['accepted']}\nEmpty: {statuses['empty']}\n"
        f"Rejected: {statuses['rejected']}\nIgnored: {statuses['ignored']}\nRows combined: {len(rows)}\n"
        f"Amount total: {business['amount_total']}\nRows are retained by source identity, including equal-looking records.\n"
        'See manifest.csv for every file and rejections.csv for reasons.\n')
    return report

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', default='input/files')
    parser.add_argument('--output', default='output')
    args = parser.parse_args()
    report = combine(args.input, args.output)
    print(f"Combined {report['output_rows']} rows; rejected {report['files']['rejected']} files. See report.json.")
