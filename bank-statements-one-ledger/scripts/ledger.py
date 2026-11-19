"""Consolidate named statement layouts; calculate in integer pennies."""
import argparse
import csv
import hashlib
import json
import re
import time
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from metadata import clean_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

SCHEMAS = {
    'layout-a.csv': ['id', 'date', 'description', 'debit', 'credit', 'balance'],
    'layout-b.xlsx': ['id', 'booking_date', 'details', 'amount', 'balance'],
    'layout-c.csv': ['reference', 'date', 'memo', 'amount', 'direction', 'balance'],
}
NAVY, TEAL, RED, PALE = '183541', '087E83', 'A43232', 'EAF3F2'


def pennies(value, blank_zero=False):
    if value is None or str(value).strip() == '':
        if blank_zero:
            return 0
        raise ValueError('missing money value')
    if isinstance(value, bool):
        raise ValueError('invalid money value')
    text = str(value).strip()
    if not re.fullmatch(r'-?\d+(?:\.\d{1,2})?', text):
        raise ValueError('money must be a plain number with at most two decimal places')
    try:
        amount = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError('invalid money value') from exc
    return int(amount * 100)


def money(value):
    return format(Decimal(value) / 100, '.2f')


def read_date(value, order):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    for fmt in ('%Y-%m-%d', '%d/%m/%Y' if order == 'DMY' else '%m/%d/%Y'):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError('invalid date for declared date order')


def table_rows(path, required):
    if path.suffix == '.xlsx':
        book = load_workbook(path, data_only=False, read_only=True)
        try:
            if 'Transactions' not in book.sheetnames:
                raise ValueError('missing Transactions sheet')
            values = list(book['Transactions'].values)
        finally:
            book.close()
    else:
        with path.open(newline='', encoding='utf-8-sig') as handle:
            values = list(csv.reader(handle))
    for index, values_row in enumerate(values):
        headers = [str(v).strip() if v is not None else '' for v in values_row]
        if set(required).issubset(headers):
            if any(headers.count(h) != 1 for h in required):
                raise ValueError('duplicate required column')
            mapping = {h: headers.index(h) for h in required}
            result = []
            for rownum, row in enumerate(values[index + 1:], index + 2):
                if not any(v is not None and str(v).strip() for v in row):
                    continue
                result.append((rownum, {h: row[i] if i < len(row) else None for h, i in mapping.items()}))
            return result
    raise ValueError('required headers not found')


def safe_text(value):
    return '' if value is None else str(value).strip()


def compute(indir):
    indir = Path(indir)
    config = json.loads((indir / 'accounts.json').read_text())
    if set(config) != set(SCHEMAS):
        raise ValueError('accounts.json must map exactly the three supported filenames')
    accounts_seen, ledger, reconciliation, rejected, issues = set(), [], [], [], []
    source_count = 0
    for filename, required in SCHEMAS.items():
        setting = config[filename]
        account = setting['account']
        if not isinstance(account, str) or not re.fullmatch(r'[A-Za-z0-9_-]+', account) or account in accounts_seen:
            raise ValueError('account codes must be distinct nonempty letters, numbers, underscores or hyphens')
        accounts_seen.add(account)
        if setting['date_order'] not in ('DMY', 'MDY'):
            raise ValueError('date_order must be DMY or MDY')
        if setting.get('currency') != 'GBP':
            raise ValueError('this case requires GBP for all accounts')
        opening, closing = pennies(setting['opening_balance']), pennies(setting['closing_balance'])
        rows = table_rows(indir / filename, required)
        source_count += len(rows)
        running, credits, debits, bad, mismatch_count, accepted = opening, 0, 0, 0, 0, 0
        last_reported = None
        seen_ids = set()
        for rownum, row in rows:
            identity = {'account': account, 'source_file': filename, 'source_row': rownum}
            try:
                id_key = 'reference' if filename == 'layout-c.csv' else 'id'
                record_id = safe_text(row[id_key])
                if not record_id:
                    raise ValueError('missing transaction identifier')
                if record_id in seen_ids:
                    raise ValueError('duplicate transaction identifier within account')
                # Reserve identifiers even if another field is invalid: never choose a correction silently.
                seen_ids.add(record_id)
                dated = read_date(row['booking_date' if filename == 'layout-b.xlsx' else 'date'], setting['date_order'])
                reported = pennies(row['balance'])
                if filename == 'layout-a.csv':
                    debit, credit = pennies(row['debit'], True), pennies(row['credit'], True)
                    if debit < 0 or credit < 0 or (debit and credit):
                        raise ValueError('debit and credit must be nonnegative; only one may be nonzero')
                    amount, description = credit - debit, safe_text(row['description'])
                elif filename == 'layout-b.xlsx':
                    amount, description = pennies(row['amount']), safe_text(row['details'])
                else:
                    amount, description = pennies(row['amount']), safe_text(row['memo'])
                    if amount < 0 or row['direction'] not in ('IN', 'OUT'):
                        raise ValueError('amount must be unsigned and direction must be IN or OUT')
                    amount *= 1 if row['direction'] == 'IN' else -1
                running += amount
                credits += max(amount, 0)
                debits += max(-amount, 0)
                accepted += 1
                last_reported = reported
                delta = reported - running if bad == 0 else None
                status = 'UNVERIFIED' if bad else ('OK' if delta == 0 else 'MISMATCH')
                if status == 'MISMATCH':
                    mismatch_count += 1
                    issues.append({**identity, 'check': 'running_balance', 'expected': money(running), 'reported': money(reported), 'difference': money(delta)})
                ledger.append({**identity, 'id': record_id, 'date': dated, 'description': description,
                               'amount': money(amount), 'reported_balance': money(reported),
                               'calculated_balance': money(running) if not bad else '',
                               'difference': money(delta) if delta is not None else '', 'check': status})
            except (ValueError, TypeError) as exc:
                bad += 1
                rejected.append({**identity, 'reason': str(exc), 'raw': row})
        if not bad and closing != running:
            issues.append({'account': account, 'source_file': 'accounts.json', 'source_row': '', 'check': 'closing_balance', 'expected': money(running), 'reported': money(closing), 'difference': money(closing - running)})
        if not bad and last_reported is not None and closing != last_reported:
            issues.append({'account': account, 'source_file': filename, 'source_row': rows[-1][0], 'check': 'last_balance_vs_declared_close', 'expected': money(closing), 'reported': money(last_reported), 'difference': money(last_reported - closing)})
        status = 'INCOMPLETE' if bad else ('OK' if mismatch_count == 0 and closing == running and (last_reported is None or last_reported == closing) else 'REVIEW')
        account_dates = [r['date'] for r in ledger if r['account'] == account]
        reconciliation.append({'period_start': min(account_dates) if account_dates else '',
                               'period_end': max(account_dates) if account_dates else '', 'account': account, 'source_file': filename, 'input_records': len(rows), 'ledger_records': accepted,
                               'rejected_records': bad, 'opening_balance': money(opening), 'credits': money(credits), 'debits': money(debits),
                               'net_movement': money(credits - debits), 'calculated_close': money(running) if not bad else '',
                               'declared_close': money(closing), 'closing_difference': money(closing - running) if not bad else '',
                               'running_mismatches': mismatch_count, 'status': status})
    report = {'input_records': source_count, 'output_records': len(ledger), 'excluded_records': len(rejected),
              'exclusion_reasons': rejected, 'currency': 'GBP', 'accounts': reconciliation,
              'running_balance_mismatches': sum(a['running_mismatches'] for a in reconciliation),
              'closing_balance_mismatches': sum(a['closing_difference'] not in ('0.00', '') for a in reconciliation),
              'total_credits': money(sum(pennies(a['credits']) for a in reconciliation)),
              'total_debits': money(sum(pennies(a['debits']) for a in reconciliation)),
              'net_movement': money(sum(pennies(a['net_movement']) for a in reconciliation)),
              'checks': {'records_conserved': source_count == len(ledger) + len(rejected),
                         'provenance_unique': len({(r['source_file'], r['source_row']) for r in ledger}) == len(ledger)},
              'result': 'INPUT_ERRORS' if rejected else ('REVIEW_REQUIRED' if issues else 'BALANCED')}
    return ledger, reconciliation, rejected, issues, report


def style_sheet(ws, widths=None):
    ws.freeze_panes = 'A2'
    ws.sheet_view.showGridLines = False
    for cell in ws[1]:
        cell.fill = PatternFill('solid', fgColor=NAVY)
        cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name='Calibri', size=11, color=NAVY)
            if cell.row % 2 == 0:
                cell.fill = PatternFill('solid', fgColor=PALE)
            # Keep imported text inert, including descriptions starting with '='.
            if isinstance(cell.value, str):
                cell.data_type = 's'
    for col in ws.columns:
        letter = col[0].column_letter
        ws.column_dimensions[letter].width = (widths or {}).get(letter, min(35, max(14, max(len(str(c.value or '')) for c in col) + 2)))
    ws.auto_filter.ref = ws.dimensions


def new_book():
    book = Workbook()
    book.remove(book.active)
    book.properties.creator = ''
    book.properties.lastModifiedBy = ''
    return book


def add_table(book, name, rows, fields):
    ws = book.create_sheet(name)
    ws.append(fields)
    for row in rows:
        ws.append([row.get(k, '') for k in fields])
    style_sheet(ws)
    # Money columns are stored as numeric cells; exact arithmetic is already complete.
    money_fields = {'amount', 'reported_balance', 'calculated_balance', 'difference', 'opening_balance', 'credits', 'debits', 'net_movement', 'calculated_close', 'declared_close', 'closing_difference', 'expected', 'reported'}
    for j, field in enumerate(fields, 1):
        if field in money_fields:
            for i in range(2, ws.max_row + 1):
                c = ws.cell(i, j)
                if c.value != '':
                    c.value = float(Decimal(str(c.value)))
                    c.number_format = '#,##0.00;[Red](#,##0.00)'
    if rows:
        tab = Table(displayName=name.replace(' ', ''), ref=ws.dimensions)
        tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(tab)
    return ws


def proof_sheet(book, report):
    ws = book.create_sheet('Balance checks', 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 17
    rows = [ ['ONE LEDGER. EVERY BALANCE CHECKED.', '', ''],
             [f"{report['input_records']} source rows → {report['output_records']} ledger rows", '', ''],
             ['Account', 'Calculated close', 'Check'],
             *[[a['account'], f"£{a['calculated_close']}" if a['calculated_close'] else 'Incomplete', a['status']] for a in report['accounts']],
             [next((f"{a['account']}: {a['closing_difference']} GBP closing difference" for a in report['accounts'] if a['closing_difference'] not in ('0.00', '')), 'No closing differences'), '', ''],
             ['Differences stay visible. No balancing entries.', '', ''] ]
    for row in rows:
        ws.append(row)
    for row in ws:
        ws.row_dimensions[row[0].row].height = 29
        for c in row:
            c.font = Font(name='Calibri', size=14, color=NAVY)
            c.alignment = Alignment(vertical='center')
    for i in (1, 2, 7, 8):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
    ws.row_dimensions[1].height = 32
    ws['A1'].font = Font(name='Calibri', size=15, bold=True, color='FFFFFF')
    for c in ws[1]: c.fill = PatternFill('solid', fgColor=NAVY)
    for c in ws[3]: c.font = Font(name='Calibri', size=12, bold=True, color=TEAL)
    for i, a in enumerate(report['accounts'], 4):
        for c in ws[i]: c.fill = PatternFill('solid', fgColor='FBE9E5' if a['status'] != 'OK' else PALE)
        ws.cell(i, 3).font = Font(name='Calibri', size=14, bold=True, color=RED if a['status'] != 'OK' else TEAL)
    ws['A7'].font = Font(name='Calibri', size=14, bold=True, color=RED if report['result'] != 'BALANCED' else TEAL)
    ws['A8'].font = Font(name='Calibri', size=12, italic=True, color=NAVY)
    ws.print_area = 'A1:C8'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A5
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_options.horizontalCentered = True
    return ws


def write_csv(path, rows, fields):
    with path.open('w', newline='', encoding='utf-8') as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction='ignore')
        writer.writeheader()
        for row in rows:
            # Safe for opening the evidence CSV directly in a spreadsheet.
            writer.writerow({k: ("'" + v if isinstance(v, str) and v.startswith(('=', '+', '@')) else v) for k, v in row.items()})


def build(indir, outdir):
    started = time.perf_counter()
    indir, outdir = Path(indir).resolve(), Path(outdir).resolve()
    if indir == outdir or indir in outdir.parents or outdir in indir.parents:
        raise ValueError('input and output directories must be separate and non-nested')
    ledger, recon, rejected, issues, report = compute(indir)
    outdir.mkdir(parents=True, exist_ok=True)
    book = new_book()
    proof_sheet(book, report)
    for name, rows, fields in (
        ('Ledger', ledger, ['account', 'source_file', 'source_row', 'id', 'date', 'description', 'amount', 'reported_balance', 'calculated_balance', 'difference', 'check']),
        ('Reconciliation', recon, list(recon[0])),
        ('Exceptions', issues, ['account', 'source_file', 'source_row', 'check', 'expected', 'reported', 'difference']),
        ('Rejected', rejected, ['account', 'source_file', 'source_row', 'reason'])):
        add_table(book, name, rows, fields)
        write_csv(outdir / (name.lower() + '.csv'), rows, fields)
    from proof import add_sources, add_trace, add_hero_before, add_hero_after
    add_trace(book, ledger)
    add_hero_after(book, ledger)
    book.save(outdir / 'ledger.xlsx')
    clean_workbook(outdir / 'ledger.xlsx')
    overview = new_book()
    for filename, fields in SCHEMAS.items():
        ws = overview.create_sheet(filename.split('.')[0])
        ws.append(['source_row'] + fields)
        for n, row in table_rows(indir / filename, fields):
            ws.append([n] + [row[f] for f in fields])
        style_sheet(ws)
    source_extracts = []
    for filename, fields in SCHEMAS.items():
        rows = table_rows(indir / filename, fields)
        if rows:
            rownum, row = rows[0]
            source_extracts.append((filename, rownum, row))
    add_sources(overview, source_extracts)
    add_hero_before(overview, table_rows(indir / 'layout-c.csv', SCHEMAS['layout-c.csv']))
    overview.save(outdir / 'input-overview.xlsx')
    clean_workbook(outdir / 'input-overview.xlsx')
    report['input_sha256'] = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in [indir / 'accounts.json', *[indir / name for name in SCHEMAS]]}
    report['runtime_seconds'] = round(time.perf_counter() - started, 6)
    (outdir / 'report.json').write_text(json.dumps(report, indent=2, default=str) + '\n')
    lines = ['# Statement reconciliation', '', f"{report['input_records']} input records; {report['output_records']} ledger records; {report['excluded_records']} rejected records.", '',
             '| Account | Opening | Credits | Debits | Calculated close | Declared close | Difference | Status |',
             '|---|---:|---:|---:|---:|---:|---:|---|']
    for a in recon:
        lines.append('| ' + ' | '.join(str(a[k]) for k in ('account', 'opening_balance', 'credits', 'debits', 'calculated_close', 'declared_close', 'closing_difference', 'status')) + ' |')
    lines += ['', 'All money is GBP. Difference = reported minus calculated. Source order is retained.', '',
              '## Exceptions', '', '| Account | Source | Row | Check | Expected | Reported | Difference |', '|---|---|---:|---|---:|---:|---:|']
    for r in issues:
        lines.append('| ' + ' | '.join(str(r[k]) for k in ('account', 'source_file', 'source_row', 'check', 'expected', 'reported', 'difference')) + ' |')
    lines += ['', 'Rejected records, if any, are in rejected.csv and report.json. An incomplete account cannot receive an OK status.', '']
    (outdir / 'reconciliation.md').write_text('\n'.join(lines))
    return report


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--input', type=Path, default=Path('input'))
    p.add_argument('--output', type=Path, default=Path('output'))
    args = p.parse_args()
    try:
        report = build(args.input, args.output)
    except (ValueError, KeyError, OSError, json.JSONDecodeError) as exc:
        print('FAIL: invalid input (' + type(exc).__name__ + '). Check the documented input schema.')
        return 2
    print(f"{report['input_records']} input rows; {report['output_records']} ledger rows; {report['excluded_records']} rejected; {report['result']}")
    return 2 if report['excluded_records'] else 0

if __name__ == '__main__':
    raise SystemExit(main())
