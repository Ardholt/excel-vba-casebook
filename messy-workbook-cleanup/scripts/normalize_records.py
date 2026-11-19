"""Deterministic, exact-match cleanup with a source-row audit."""
import argparse
import csv
import hashlib
import json
import re
import time
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import openpyxl
from workbook_meta import scrub_properties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FIELDS = ('row_id', 'customer_id', 'name', 'date', 'amount')


def clean_text(value):
    return ' '.join(str(value).split()) if value is not None else ''


def parse_date(value, order):
    if isinstance(value, datetime):
        if value.time().isoformat() != '00:00:00':
            raise ValueError('date contains a time')
        return value.date()
    if isinstance(value, date):
        return value
    value = clean_text(value)
    try:
        if re.fullmatch(r'\d{4}-\d{2}-\d{2}', value):
            return date.fromisoformat(value)
        if re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}', value):
            a, b, y = map(int, value.split('/'))
            return date(y, b, a) if order == 'DMY' else date(y, a, b)
    except ValueError:
        pass
    raise ValueError('invalid date; use ISO or slash date under ' + order)


def parse_amount(value):
    if isinstance(value, bool) or value is None:
        raise ValueError('invalid amount')
    s = clean_text(value)
    # Decimal point, correctly grouped commas, optional currency symbol and sign.
    parentheses = s.startswith('(') and s.endswith(')')
    if parentheses:
        s = s[1:-1].strip()
    m = re.fullmatch(r'([+-]?)([£$€]?)\s*([+-]?)([0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(\.[0-9]{1,2})?\s*([£$€]?)', s)
    if not m or (m[1] and m[3]) or (m[2] and m[6]) or (parentheses and (m[1] or m[3])):
        raise ValueError('invalid amount; use decimal point and up to two decimal places')
    amount = Decimal(m[4].replace(',', '') + (m[5] or ''))
    if parentheses or (m[1] or m[3]) == '-':
        amount = -amount
    # Keep cents exact in workbook numeric storage.
    if abs(amount) > Decimal('999999999999.99'):
        raise ValueError('amount exceeds supported precision')
    return amount.quantize(Decimal('0.01'))


def visible(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return '' if value is None else str(value)


def literal(cell, value):
    cell.value = value
    if isinstance(value, str):
        cell.data_type = 's'


def style(sheet, widths=None):
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_view.showGridLines = False
    for cell in sheet[1]:
        cell.fill = PatternFill('solid', fgColor='183E45')
        cell.font = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name='Calibri', size=11, color='16333A')
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            if cell.row % 2 == 0:
                cell.fill = PatternFill('solid', fgColor='EDF5F3')
        sheet.row_dimensions[row[0].row].height = 27
    for i in range(1, sheet.max_column + 1):
        sheet.column_dimensions[get_column_letter(i)].width = (widths or {}).get(i, 23)


def table(book, title, headers, rows):
    ws = book.create_sheet(title)
    ws.append(list(headers))
    for values in rows:
        for col, value in enumerate(values, 1):
            literal(ws.cell(ws.max_row + 1 if col == 1 else ws.max_row, col), value)
    style(ws)
    return ws


def write_csv(path, headers, rows):
    # JSON-style escaping is unnecessary: quoting follows RFC 4180. Read as data.
    with path.open('w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


def load_records(path):
    book = openpyxl.load_workbook(path, data_only=False)
    if 'Records' not in book.sheetnames:
        raise ValueError('missing Records sheet')
    ws = book['Records']
    header_row = None
    for row in ws.iter_rows():
        labels = [clean_text(c.value) for c in row]
        if set(FIELDS).issubset(labels):
            if any(labels.count(f) != 1 for f in FIELDS):
                raise ValueError('duplicate required header')
            header_row = row[0].row
            columns = {f: labels.index(f) + 1 for f in FIELDS}
            break
    if header_row is None:
        raise ValueError('missing required column headers')
    result = []
    for row in ws.iter_rows(min_row=header_row + 1):
        if all(c.value is None or clean_text(c.value) == '' for c in row):
            continue
        cells = {f: ws.cell(row[0].row, c) for f, c in columns.items()}
        result.append((row[0].row, cells))
    return result


def run(input_dir, output_dir):
    started = time.perf_counter()
    input_dir, output_dir = Path(input_dir).resolve(), Path(output_dir).resolve()
    if input_dir == output_dir or input_dir in output_dir.parents or output_dir in input_dir.parents:
        raise ValueError('input and output directories must be separate, non-nested folders')
    rules = json.loads((input_dir / 'rules.json').read_text())
    order = rules.get('date_order')
    if order not in ('DMY', 'MDY'):
        raise ValueError('date_order must be DMY or MDY')
    records = load_records(input_dir / 'input.xlsx')
    kept, duplicates, rejected, audit, changes = [], [], [], [], []
    seen = {}
    valid_total = Decimal('0.00')
    duplicate_total = Decimal('0.00')
    for source_row, cells in records:
        raw = {f: c.value for f, c in cells.items()}
        normalized = {f: (visible(raw[f]).strip() if f != 'name' else clean_text(raw[f])) for f in FIELDS[:3]}
        reasons = []
        for f in FIELDS[:3]:
            if not normalized[f]:
                reasons.append(f + ' is blank')
        # Numeric IDs cannot reveal lost leading zeros; require text rather than guessing.
        for f in ('row_id', 'customer_id'):
            if raw[f] is not None and not isinstance(raw[f], str):
                reasons.append(f + ' must be stored as text')
        if any(c.data_type == 'f' for c in cells.values()):
            reasons.append('formula input is not accepted')
        for field, parser in [('date', lambda v: parse_date(v, order)), ('amount', parse_amount)]:
            try:
                normalized[field] = parser(raw[field])
            except (ValueError, InvalidOperation):
                reasons.append('invalid ' + field)
        raw_values = [visible(raw[f]) for f in FIELDS]
        if reasons:
            reason = '; '.join(reasons)
            rejected.append([source_row] + raw_values + [reason])
            audit.append([source_row, normalized['row_id'], 'rejected', '', reason] + raw_values)
            continue
        amount = normalized['amount']
        valid_total += amount
        key = (normalized['customer_id'], normalized['name'].casefold(), normalized['date'], amount)
        status, target, reason = 'kept', '', 'valid unique record'
        if key in seen:
            first = seen[key]
            status, target, reason = 'duplicate', first[0], 'exact normalized match; first source row retained'
            duplicate_total += amount
            duplicates.append([source_row, normalized['row_id'], first[0], first[1], format(amount, '.2f')])
        else:
            seen[key] = (source_row, normalized['row_id'])
            kept.append([normalized[f] for f in FIELDS] + [source_row])
        after = [visible(normalized[f]) for f in FIELDS]
        for f, old, new in zip(FIELDS, raw_values, after):
            if f == 'date' and isinstance(raw[f], (date, datetime)):
                old = normalized[f].isoformat()
            if f == 'amount' and isinstance(raw[f], (int, float, Decimal)):
                old = format(normalized[f], '.2f')
            if old != new:
                changes.append([source_row, normalized['row_id'], f, old, new, status])
        audit.append([source_row, normalized['row_id'], status, target, reason] + after)
    output_dir.mkdir(parents=True, exist_ok=True)
    book = openpyxl.Workbook()
    book.remove(book.active)
    book.properties.creator = book.properties.lastModifiedBy = 'Genco Doganay'
    ws = table(book, 'Records', list(FIELDS) + ['source_row'],
               [[*r[:3], datetime.combine(r[3], datetime.min.time()), float(r[4]), r[5]] for r in kept])
    for row in ws.iter_rows(min_row=2):
        row[3].number_format = 'yyyy-mm-dd'
        row[4].number_format = '#,##0.00;[Red](#,##0.00)'
        row[1].number_format = '@'
    duplicate_headers = ['source_row', 'row_id', 'kept_source_row', 'kept_row_id', 'amount']
    reject_headers = ['source_row', *FIELDS, 'reason']
    audit_headers = ['source_row', 'row_id', 'status', 'kept_source_row', 'reason', *['value_' + f for f in FIELDS]]
    change_headers = ['source_row', 'row_id', 'field', 'before', 'after', 'status']
    table(book, 'Duplicates', duplicate_headers, duplicates)
    table(book, 'Rejected', reject_headers, rejected)
    table(book, 'Audit', audit_headers, audit)
    table(book, 'Changes', change_headers, changes)
    book.save(output_dir / 'customer-register.xlsx')
    scrub_properties(output_dir / 'customer-register.xlsx')
    for name, headers, rows in [('duplicates', duplicate_headers, duplicates), ('rejected', reject_headers, rejected),
                                ('audit', audit_headers, audit), ('changes', change_headers, changes)]:
        write_csv(output_dir / (name + '.csv'), headers, rows)
    kept_total = sum((r[4] for r in kept), Decimal('0.00'))
    canonical = [[visible(v) for v in r[:5]] for r in kept]
    report = {
        'date_order': order, 'input_records': len(records), 'kept_records': len(kept),
        'duplicate_records': len(duplicates), 'rejected_records': len(rejected),
        'field_changes': len(changes), 'valid_source_amount': format(valid_total, '.2f'),
        'kept_amount': format(kept_total, '.2f'), 'duplicate_amount': format(duplicate_total, '.2f'),
        'rejected_amount_policy': 'Excluded from totals; raw values and reasons retained in rejected.csv.',
        'checks': {'every_source_row_accounted_for': len(records) == len(audit) == len(kept) + len(duplicates) + len(rejected),
                   'valid_amount_conserved': valid_total == kept_total + duplicate_total,
                   'unique_source_rows': len({r[0] for r in audit}) == len(records)},
        'business_result_sha256': hashlib.sha256(json.dumps(canonical, ensure_ascii=False).encode()).hexdigest(),
        'inputs': {f: hashlib.sha256((input_dir / f).read_bytes()).hexdigest() for f in ('input.xlsx', 'rules.json')},
        'runtime_seconds': round(time.perf_counter() - started, 6),
        'software': {'openpyxl': openpyxl.__version__}}
    (output_dir / 'report.json').write_text(json.dumps(report, indent=2) + '\n')
    (output_dir / 'summary.md').write_text(
        '# Cleanup result\n\n'
        f"{len(records)} original entries = {len(kept)} kept + {len(duplicates)} removed repeats + {len(rejected)} saved for review.\n\n"
        f"Valid source amount {valid_total:.2f} = kept {kept_total:.2f} + removed repeats {duplicate_total:.2f}.\n\n"
        'Rejected records remain outside the amount reconciliation. See rejected.csv for raw values and reasons.\n\n'
        'Counts and checks: [report.json](report.json). Row decisions: [audit.csv](audit.csv).\n')
    if not all(report['checks'].values()):
        raise ValueError('record or amount reconciliation failed')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', default='input')
    parser.add_argument('--output', default='output')
    args = parser.parse_args()
    try:
        result = run(args.input, args.output)
        print(f"{result['input_records']} source = {result['kept_records']} kept + {result['duplicate_records']} duplicates + {result['rejected_records']} rejected")
    except (ValueError, OSError, KeyError, json.JSONDecodeError) as exc:
        raise SystemExit('FAIL: ' + str(exc))
