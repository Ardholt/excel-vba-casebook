"""Propose similar records; change identity only from an explicit confirmation file."""
import argparse
import csv
import hashlib
import json
import re
import time
import unicodedata
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path
import openpyxl
from metadata import clean_properties
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HEADERS = ['customer_id', 'name', 'postcode']


def normal(value):
    return ' '.join(unicodedata.normalize('NFKC', str(value or '')).casefold().split())


def workbook(path, sheets):
    book = openpyxl.Workbook()
    book.remove(book.active)
    book.properties.creator = 'Genco Doganay'
    book.properties.lastModifiedBy = 'Genco Doganay'
    for title, headers, rows in sheets:
        sheet = book.create_sheet(title)
        sheet.append(headers)
        for row in rows:
            sheet.append(list(row))
        for row in sheet:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = 's'
                cell.font = Font(name='Carlito', size=12)
                cell.alignment = Alignment(vertical='center')
        for cell in sheet[1]:
            cell.fill = PatternFill('solid', fgColor='183D4A')
            cell.font = Font(name='Carlito', size=12, color='FFFFFF', bold=True)
        for row in range(2, sheet.max_row + 1):
            sheet.row_dimensions[row].height = 25
            if row % 2 == 0:
                for cell in sheet[row]:
                    cell.fill = PatternFill('solid', fgColor='EFF6F5')
        sheet.row_dimensions[1].height = 28
        for col, header in enumerate(headers, 1):
            width = min(55, max(len(str(header)), *(len(str(r[col-1] or '')) for r in rows), 10) + 3) if rows else len(header)+4
            sheet.column_dimensions[get_column_letter(col)].width = width
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.orientation = 'landscape'
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 1
        sheet.print_options.horizontalCentered = True
        sheet.print_area = sheet.dimensions
        sheet.oddFooter.center.text = ''
    book.save(path)
    clean_properties(path)


def read_customers(path):
    book = openpyxl.load_workbook(path, data_only=False)
    if 'Customers' not in book.sheetnames:
        raise ValueError('Missing Customers sheet')
    sheet = book['Customers']
    header = None
    result = []
    seen = set()
    for rowno, row in enumerate(sheet.iter_rows(), 1):
        vals = [str(c.value).strip() if c.value is not None else '' for c in row]
        if header is None:
            if all(h in vals for h in HEADERS):
                if any(vals.count(h) != 1 for h in HEADERS):
                    raise ValueError('Repeated required header')
                header = {h: vals.index(h) for h in HEADERS}
            continue
        if not any(vals):
            continue
        record = {h: vals[i] for h, i in header.items()}
        if any(row[i].data_type == 'f' for i in header.values()):
            raise ValueError('Customer fields must be values, not formulas')
        key = record['customer_id']
        if not key or key in seen:
            raise ValueError('Customer IDs must be nonblank and unique')
        if not isinstance(row[header['customer_id']].value, str):
            raise ValueError('Customer IDs must be text to preserve leading zeros')
        if not record['name']:
            raise ValueError('Customer names must be nonblank')
        record['source_row'] = rowno
        result.append(record)
        seen.add(key)
    if header is None:
        raise ValueError('Missing required customer headers')
    return result


def propose(records):
    blocks = defaultdict(list)
    for row in records:
        if normal(row['postcode']):
            blocks[normal(row['postcode'])].append(row)
    result = []
    for area in sorted(blocks):
        group = sorted(blocks[area], key=lambda r: r['customer_id'])
        for i, left in enumerate(group):
            for right in group[i+1:]:
                a, b = normal(left['name']), normal(right['name'])
                score = SequenceMatcher(None, a, b, autojunk=False).ratio()
                if score >= .82:
                    result.append([left['customer_id'], right['customer_id'], left['name'], right['name'], round(score, 4), 'Same area; normalized names equal' if a == b else 'Same area; close name spelling', 'Review only'])
    return result


def read_confirmed(path):
    if not path.exists():
        return []
    with path.open(newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None or any(reader.fieldnames.count(h) != 1 for h in ['source_id', 'target_id']):
            raise ValueError('Confirmation headers must include source_id,target_id once')
        return [(str(r.get('source_id') or '').strip(), str(r.get('target_id') or '').strip()) for r in reader if any(r.values())]


def resolve(records, confirmations):
    ids = {r['customer_id'] for r in records}
    links = {}
    errors = []
    for line, (source, target) in enumerate(confirmations, 2):
        if source not in ids or target not in ids:
            errors.append([line, source, target, 'Unknown or blank customer ID'])
        elif source == target:
            errors.append([line, source, target, 'Self mapping is not a merge'])
        elif source in links and links[source] != target:
            errors.append([line, source, target, 'Conflicting targets for source'])
        else:
            links[source] = target
    for source in links:
        trail, current = set(), source
        while current in links:
            if current in trail:
                errors.append(['', source, links[source], 'Cycle in confirmations'])
                break
            trail.add(current)
            current = links[current]
    if errors:
        return {i: i for i in ids}, {i: i for i in ids}, errors
    resolved, paths = {}, {}
    for source in ids:
        chain = [source]
        while chain[-1] in links:
            chain.append(links[chain[-1]])
        resolved[source] = chain[-1]
        paths[source] = ' > '.join(chain)
    return resolved, paths, []


def write_csv(path, headers, rows):
    # Prefix active spreadsheet expressions when CSVs are opened interactively.
    def safe(v):
        return "'"+v if isinstance(v, str) and v.startswith(('=', '+', '-', '@', '\t', '\r')) else v
    with path.open('w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows([[safe(v) for v in row] for row in rows])


def run(input_dir, output_dir):
    started = time.perf_counter()
    input_dir, output_dir = Path(input_dir).resolve(), Path(output_dir).resolve()
    if input_dir == output_dir or input_dir in output_dir.parents or output_dir in input_dir.parents:
        raise ValueError('Input and output folders must be separate')
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / 'report.json').write_text(json.dumps({'status':'INCOMPLETE','outputs_current':False})+'\n')
    try:
        records = read_customers(input_dir / 'input.xlsx')
        confirmations = read_confirmed(input_dir / 'confirmed.csv')
    except Exception as error:
        (output_dir / 'report.json').write_text(json.dumps({'status':'FAIL', 'reason':str(error) if isinstance(error, ValueError) else 'Required input could not be read', 'outputs_current':False}, indent=2)+'\n')
        if isinstance(error, ValueError):
            raise
        raise ValueError('Required input could not be read: '+type(error).__name__) from None
    proposals = propose(records)
    mapping, paths, errors = resolve(records, confirmations)
    retained = [r for r in records if mapping[r['customer_id']] == r['customer_id']]
    proposed_ids = {p[j] for p in proposals for j in (0, 1)}
    audit = []
    for r in records:
        source = r['customer_id']
        reason = 'Explicit confirmation' if mapping[source] != source else ('Batch rejected; unchanged' if errors else ('Not confirmed; unchanged' if source in proposed_ids else 'No confirmation; unchanged'))
        audit.append([source, mapping[source], paths[source], reason, r['source_row'], r['name'], r['postcode']])
    ph = ['left_id', 'right_id', 'left_name', 'right_name', 'name_similarity', 'reason', 'status']
    ah = ['source_id', 'retained_id', 'confirmation_path', 'reason', 'source_row', 'original_name', 'original_postcode']
    eh = ['csv_row', 'source_id', 'target_id', 'reason']
    customer_rows = [[r[h] for h in HEADERS] for r in retained]
    workbook(output_dir / 'result.xlsx', [('Customers', HEADERS, customer_rows), ('Proposed', ph, proposals), ('Confirmed', ['source_id', 'target_id'], confirmations), ('Audit', ah, audit), ('Errors', eh, errors)])
    workbook(output_dir / 'before.xlsx', [('Customers', HEADERS, [[r[h] for h in HEADERS] for r in records])])
    write_csv(output_dir / 'proposed.csv', ph, proposals)
    write_csv(output_dir / 'audit.csv', ah, audit)
    write_csv(output_dir / 'confirmation-errors.csv', eh, errors)
    write_csv(output_dir / 'confirmed-template.csv', ['source_id', 'target_id'], [])
    merged = sum(s != t for s, t in mapping.items())
    report = {'status': 'REJECTED' if errors else 'PASS', 'input_records': len(records), 'retained_records': len(retained), 'merged_records': merged, 'excluded_records': 0, 'proposed_pairs': len(proposals), 'confirmation_rows': len(confirmations), 'confirmation_errors': len(errors), 'unconfirmed_proposed_pairs': sum(mapping[p[0]] != mapping[p[1]] for p in proposals), 'unchanged_without_confirmations': len(records), 'audit_rows': len(audit), 'conservation': len(records) == len(retained)+merged == len(audit), 'runtime_seconds': round(time.perf_counter()-started, 4), 'input_sha256': hashlib.sha256((input_dir/'input.xlsx').read_bytes()).hexdigest(), 'confirmed_sha256': hashlib.sha256((input_dir/'confirmed.csv').read_bytes()).hexdigest() if (input_dir/'confirmed.csv').exists() else None, 'policy': 'Only explicit confirmations change identity. Any invalid confirmation rejects the entire batch. Retained target fields are preserved; source fields remain in Audit.'}
    (output_dir / 'report.json').write_text(json.dumps(report, indent=2)+'\n')
    labels = {'input_records':'Starting customer entries', 'retained_records':'Finished customer entries', 'merged_records':'Approved joins', 'proposed_pairs':'Suggestions to review', 'unconfirmed_proposed_pairs':'Suggested pairs left separate', 'confirmation_errors':'Approval errors', 'audit_rows':'Original entries recorded'}
    (output_dir / 'evidence.md').write_text(f"# Customer review result\n\nStatus: {report['status']}.\n\n| Check | Result |\n|---|---:|\n" + ''.join(f"| {label} | {report[key]} |\n" for key,label in labels.items()) + '\nEvery original entry is listed in [the complete record](audit.csv). Suggested pairs stay separate unless approved. See [report.json](report.json) for timing and input checks.\n')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', default='input')
    parser.add_argument('--output', default='output')
    args = parser.parse_args()
    try:
        report = run(args.input, args.output)
        print(json.dumps(report, indent=2))
        raise SystemExit(0 if report['status'] == 'PASS' else 2)
    except ValueError as error:
        print('FAIL: '+str(error))
        raise SystemExit(2)
