"""Exact-decimal sales reference and formula-bearing dashboard builder."""
import argparse
import csv
import hashlib
import json
import re
from collections import Counter, defaultdict
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
import openpyxl
import xlsxwriter
from workbook_meta import scrub

HEADERS = ['sale_id', 'date', 'category', 'amount', 'status']
MONEY = '£#,##0.00;[Red]-£#,##0.00'

def clean(value):
    return '' if value is None else str(value).strip()

def read_sales(path):
    book = openpyxl.load_workbook(path, data_only=False)
    if 'Sales' not in book.sheetnames:
        raise ValueError('Missing Sales sheet')
    sheet = book['Sales']
    for cells in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 50)):
        values = [clean(c.value) for c in cells]
        if all(name in values for name in HEADERS):
            if any(values.count(name) != 1 for name in HEADERS):
                raise ValueError('Repeated required header')
            positions = [values.index(name) for name in HEADERS]
            header_row = cells[0].row
            break
    else:
        raise ValueError('Required headers not found in first 50 rows')
    rows = []
    for cells in sheet.iter_rows(min_row=header_row + 1):
        selected = [cells[i] for i in positions]
        if all(clean(c.value) == '' for c in selected):
            continue
        rows.append({'source_row': cells[0].row, **dict(zip(HEADERS, [c.value for c in selected])),
                     'formula': any(c.data_type == 'f' for c in selected),
                     'cell_error': any(c.data_type == 'e' for c in selected)})
    book.close()
    return rows

def analyze(rows):
    ids = Counter(clean(r['sale_id']) for r in rows if clean(r['sale_id']))
    accepted, audit = [], []
    for raw in rows:
        r = {k: clean(raw[k]) for k in HEADERS}
        reason = ''
        if not r['sale_id']:
            reason = 'Missing sale ID'
        elif ids[r['sale_id']] > 1:
            reason = 'Duplicate sale ID: all occurrences excluded'
        elif raw.get('formula'):
            reason = 'Formula in source row: use literal values'
        elif raw.get('cell_error'):
            reason = 'Invalid cell value'
        elif not re.fullmatch(r'[0-9]{4}-[0-9]{2}-[0-9]{2}', r['date']):
            reason = 'Invalid ISO date'
        else:
            try:
                d = date.fromisoformat(r['date'])
                if d.year < 1900:
                    raise ValueError()
            except ValueError:
                reason = 'Invalid ISO date'
        if not reason and not r['category']:
            reason = 'Missing category'
        if not reason:
            try:
                if isinstance(raw['amount'], bool) or not re.fullmatch(r'-?[0-9]+(?:\.[0-9]{1,2})?', r['amount']):
                    raise InvalidOperation()
                amount = Decimal(r['amount'])
                if abs(amount) > Decimal('999999999'):
                    raise InvalidOperation()
            except InvalidOperation:
                reason = 'Invalid amount: decimal with up to two places required'
        if not reason and r['status'] not in ('Paid', 'Pending', 'Cancelled'):
            reason = 'Invalid status'
        disposition = 'Rejected' if reason else r['status']
        audit.append({'source_row': raw['source_row'], **r, 'disposition': disposition, 'reason': reason})
        if not reason:
            accepted.append({'source_row': raw['source_row'], **r, 'amount': amount, 'month': r['date'][:7]})
    counts = Counter(r['disposition'] for r in audit)
    totals = {s: sum((r['amount'] for r in accepted if r['status'] == s), Decimal('0')) for s in ('Paid', 'Pending', 'Cancelled')}
    monthly, categories = defaultdict(Decimal), defaultdict(Decimal)
    for r in accepted:
        if r['status'] == 'Paid':
            monthly[r['month']] += r['amount']
            categories[r['category']] += r['amount']
    report = {'input_records': len(rows), 'accepted_records': len(accepted),
              'rejected_records': counts['Rejected'], 'paid_count': counts['Paid'],
              'pending_count': counts['Pending'], 'cancelled_count': counts['Cancelled'],
              'paid_revenue': f"{totals['Paid']:.2f}", 'pending_amount': f"{totals['Pending']:.2f}",
              'cancelled_amount_excluded': f"{totals['Cancelled']:.2f}",
              'negative_paid_count': sum(r['amount'] < 0 for r in accepted if r['status'] == 'Paid'),
              'duplicate_rows_excluded': sum('Duplicate' in r['reason'] for r in audit),
              'rejection_reasons': dict(Counter(r['reason'] for r in audit if r['reason'])),
              'paid_by_month': {k: f'{v:.2f}' for k, v in sorted(monthly.items())},
              'paid_by_category': {k: f'{v:.2f}' for k, v in sorted(categories.items())},
              'checks': {'records_conserved': len(rows) == sum(counts.values()),
                         'month_total_matches': sum(monthly.values()) == totals['Paid'],
                         'category_total_matches': sum(categories.values()) == totals['Paid'],
                         'accepted_ids_unique': len({r['sale_id'] for r in accepted}) == len(accepted)}}
    return accepted, audit, report

def make_book(path, source, accepted, audit, report, phase='REFRESHED', vba_project=None):
    with xlsxwriter.Workbook(path, {'strings_to_formulas': False, 'strings_to_urls': False}) as book:
        book.set_properties({'author': '', 'company': '', 'comments': ''})
        if vba_project:
            book.add_vba_project(str(vba_project))
            book.set_vba_name('ThisWorkbook')
        navy, teal, light = '#142F43', '#087F73', '#E8F3F1'
        title = book.add_format({'font_name': 'Carlito', 'font_size': 22, 'bold': True, 'font_color': navy})
        label = book.add_format({'font_name': 'Carlito', 'font_size': 14, 'font_color': navy})
        note = book.add_format({'font_name': 'Carlito', 'font_size': 12, 'font_color': '#496273'})
        head = book.add_format({'font_name': 'Carlito', 'font_size': 14, 'bold': True, 'bg_color': navy, 'font_color': 'white'})
        money = book.add_format({'font_name': 'Carlito', 'font_size': 15, 'num_format': MONEY, 'font_color': navy})
        big = book.add_format({'font_name': 'Carlito', 'font_size': 30, 'num_format': MONEY, 'bold': True, 'font_color': teal, 'bg_color': light})
        integer = book.add_format({'font_name': 'Carlito', 'font_size': 15, 'num_format': '0', 'font_color': navy})
        dash = book.add_worksheet('Dashboard')
        if vba_project:
            dash.set_vba_name('Sheet1')
            dash.insert_button('A5', {'macro': 'RefreshDashboard', 'caption': 'Refresh sales', 'width': 170, 'height': 28, 'y_offset': 2})
        dash.hide_gridlines(2)
        dash.set_column('A:A', 29)
        dash.set_column('B:B', 21)
        dash.set_column('D:H', 14)
        dash.set_default_row(25)
        dash.merge_range('A1:B1', 'SALES / ' + phase, title)
        dash.merge_range('A2:B2', 'Paid revenue · GBP · returns included', note)
        end = max(2, len(accepted) + 1)
        amounts, statuses = f'ValidSales!$E$2:$E${end}', f'ValidSales!$F$2:$F${end}'
        dash.merge_range('A3:B4', '', big)
        dash.write_formula('A3', f'=SUMIF({statuses},"Paid",{amounts})', big, float(report['paid_revenue']))
        dash.write('A6', 'Paid transactions', label)
        dash.write_formula('B6', f'=COUNTIF({statuses},"Paid")', integer, report['paid_count'])
        dash.write('A7', 'Pending · kept separate', label)
        dash.write_formula('B7', f'=SUMIF({statuses},"Pending",{amounts})', money, float(report['pending_amount']))
        dash.write('A8', 'Cancelled · excluded', label)
        dash.write('B8', report['cancelled_count'], integer)
        dash.write('A9', 'Rows to review', label)
        dash.write('B9', report['rejected_records'], integer)
        dash.write_row('A11', ['Paid by month', 'Revenue'], head)
        for row, (month, amount) in enumerate(report['paid_by_month'].items(), 11):
            dash.write(row, 0, month, label)
            dash.write_formula(row, 1, f'=SUMIFS({amounts},{statuses},"Paid",ValidSales!$G$2:$G${end},A{row+1})', money, float(amount))
        catrow = 12 + max(1, len(report['paid_by_month']))
        dash.write_row(catrow, 0, ['Paid by category', 'Revenue'], head)
        for row, (category, amount) in enumerate(report['paid_by_category'].items(), catrow + 1):
            dash.write(row, 0, category, label)
            dash.write_formula(row, 1, f'=SUMPRODUCT(--EXACT(ValidSales!$D$2:$D${end},A{row+1}),--({statuses}="Paid"),{amounts})', money, float(amount))
        for start, count, name, anchor in [(11, len(report['paid_by_month']), 'Paid by month', 'D3'),
                                          (catrow+1, len(report['paid_by_category']), 'Paid by category', 'D20')]:
            if count:
                chart = book.add_chart({'type': 'column'})
                chart.add_series({'name': name, 'categories': ['Dashboard', start, 0, start+count-1, 0],
                                  'values': ['Dashboard', start, 1, start+count-1, 1], 'fill': {'color': teal}})
                chart.set_title({'name': name})
                chart.set_legend({'none': True})
                chart.set_y_axis({'num_format': '£#,##0'})
                dash.insert_chart(anchor, chart)
        dash.merge_range(catrow+len(report['paid_by_category'])+2, 0, catrow+len(report['paid_by_category'])+2, 1, 'Reference result · see Audit for exclusions', note)
        dash.print_area(0, 0, 9, 1)
        dash.set_paper(9)
        dash.set_margins(0.15, 0.15, 0.15, 0.15)
        dash.fit_to_pages(1, 1)
        dash.set_header('')
        dash.set_footer('')
        tables = [('Sales', HEADERS, source),
                  ('ValidSales', ['source_row']+HEADERS+['month'], accepted),
                  ('Audit', ['source_row']+HEADERS+['disposition', 'reason'], audit)]
        for sheet_index, (name, headers, rows) in enumerate(tables, 2):
            sheet = book.add_worksheet(name)
            if vba_project:
                sheet.set_vba_name(f'Sheet{sheet_index}')
            sheet.write_row(0, 0, headers, head)
            for i, row in enumerate(rows, 1):
                for j, key in enumerate(headers):
                    value = row[key]
                    if isinstance(value, Decimal):
                        sheet.write_number(i, j, float(value), money)
                    elif value is not None:
                        sheet.write(i, j, value)
            sheet.set_column(0, len(headers)-1, 18)
            if name == 'Audit':
                sheet.set_column(len(headers)-1, len(headers)-1, 60)
            sheet.freeze_panes(1, 0)
            sheet.autofilter(0, 0, len(rows), len(headers)-1)
    scrub(path)

def build(input_dir, output_dir):
    src, dest = Path(input_dir), Path(output_dir)
    if src.resolve() == dest.resolve() or src.resolve() in dest.resolve().parents or dest.resolve() in src.resolve().parents:
        raise ValueError('Input and output must be separate directories')
    source = src / 'input.xlsx'
    rows = read_sales(source)
    accepted, audit, report = analyze(rows)
    dest.mkdir(parents=True, exist_ok=True)
    make_book(dest / 'dashboard.xlsx', rows, accepted, audit, report)
    report['input_sha256'] = hashlib.sha256(source.read_bytes()).hexdigest()
    before = src / 'before.xlsx'
    if before.exists():
        old = read_sales(before)
        a, b, baseline = analyze(old)
        make_book(dest / 'before-dashboard.xlsx', old, a, b, baseline, 'BEFORE REFRESH')
        report['before'] = baseline
        report['paid_revenue_change'] = f"{Decimal(report['paid_revenue']) - Decimal(baseline['paid_revenue']):.2f}"
    with (dest / 'audit.csv').open('w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['source_row']+HEADERS+['disposition','reason'])
        writer.writeheader()
        # Keep spreadsheet applications from interpreting audit text as formulas.
        def csv_literal(value):
            if isinstance(value, str) and value.startswith(('=', '+', '-', '@')) and not re.fullmatch(r'-?[0-9]+(?:\.[0-9]{1,2})?', value):
                return "'" + value
            return value
        writer.writerows({key: csv_literal(value) for key, value in row.items()} for row in audit)
    (dest / 'report.json').write_text(json.dumps(report, indent=2) + '\n')
    return report

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', default='input')
    parser.add_argument('--output', default='output')
    args = parser.parse_args()
    build(args.input, args.output)
