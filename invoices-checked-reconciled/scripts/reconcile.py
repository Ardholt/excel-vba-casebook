"""Invoice arithmetic and payment allocation, using decimal cents throughout."""
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import csv
import html
import json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

CENT = Decimal('0.01')
ZERO = Decimal('0.00')
HEADERS = {
    'Lines': ['invoice_id', 'line_id', 'quantity', 'unit_price', 'tax_rate'],
    'Invoices': ['invoice_id', 'declared_total'],
    'Payments': ['payment_id', 'invoice_id', 'amount'],
}


def decimal(value, cents=False):
    if value is None or isinstance(value, bool):
        raise ValueError('Missing or invalid number')
    try:
        result = Decimal(str(value).strip())
        if not result.is_finite() or abs(result) > Decimal('1000000000'):
            raise ValueError('Number outside supported range')
        if cents and result != result.quantize(CENT):
            raise ValueError('Amount has more than two decimal places')
        return result
    except (InvalidOperation, TypeError):
        raise ValueError('Invalid number') from None


def rounded(value):
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def money(value):
    return format(rounded(value), '.2f')


def identifier(value):
    result = '' if value is None else str(value).strip()
    if not result or len(result) > 80:
        raise ValueError('Missing or overlong identifier')
    return result


def records(path):
    book = load_workbook(path, data_only=False)
    result = {}
    for name, wanted in HEADERS.items():
        if name not in book:
            raise ValueError(f'Missing sheet: {name}')
        sheet = book[name]
        header = None
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 25)):
            values = [str(c.value).strip() if c.value is not None else '' for c in row]
            if set(wanted).issubset(values):
                if any(values.count(k) != 1 for k in wanted):
                    raise ValueError(f'Repeated header: {name}')
                header = row[0].row
                indices = {k: values.index(k) for k in wanted}
                break
        if header is None:
            raise ValueError(f'Missing required headers: {name}')
        result[name] = []
        for row in sheet.iter_rows(min_row=header + 1):
            if all(c.value is None for c in row):
                continue
            item = {k: row[i].value for k, i in indices.items()}
            item['source_row'] = row[0].row
            item['_formula'] = any(row[i].data_type == 'f' for i in indices.values())
            result[name].append(item)
    book.close()
    return result


def reconcile(source):
    raw = records(source)
    issues, line_rows, invoice_rows, payment_rows = [], [], [], []
    def issue(sheet, row, code, detail):
        issues.append({'sheet': sheet, 'source_row': row['source_row'],
                       'invoice_id': str(row.get('invoice_id') or ''),
                       'record_id': str(row.get('payment_id', row.get('line_id', row.get('invoice_id'))) or ''),
                       'reason': code, 'detail': detail})
    invoice_counts = Counter(str(r['invoice_id']).strip() for r in raw['Invoices'])
    invoices = {}
    for row in raw['Invoices']:
        try:
            iid = identifier(row['invoice_id'])
            if row['_formula']:
                raise ValueError('Input formulas are not accepted')
            if invoice_counts[iid] != 1:
                raise ValueError('Repeated invoice ID; every occurrence held')
            declared = decimal(row['declared_total'], cents=True)
            invoices[iid] = {'invoice_id': iid, 'declared_total': declared, 'source_row': row['source_row']}
        except ValueError as err:
            issue('Invoices', row, 'invalid_invoice', str(err))
    line_counts = Counter((str(r['invoice_id']).strip(), str(r['line_id']).strip()) for r in raw['Lines'])
    held_invoices = set()
    for row in raw['Lines']:
        iid = str(row['invoice_id'] or '').strip()
        item = {'invoice_id': iid, 'line_id': str(row['line_id'] or ''), 'source_row': row['source_row'],
                'quantity': '', 'unit_price': '', 'tax_rate': '', 'net': '', 'tax': '', 'gross': '', 'status': 'held'}
        try:
            iid, lid = identifier(row['invoice_id']), identifier(row['line_id'])
            if row['_formula']:
                raise ValueError('Input formulas are not accepted')
            if iid not in invoices:
                raise ValueError('No unique valid invoice header')
            if line_counts[(iid, lid)] != 1:
                raise ValueError('Repeated line ID within invoice; every occurrence held')
            qty, price, rate = (decimal(row[k]) for k in ('quantity', 'unit_price', 'tax_rate'))
            if price < 0 or not ZERO <= rate <= 1:
                raise ValueError('Unit price must be nonnegative; tax rate must be between zero and one')
            net = rounded(qty * price)
            tax = rounded(net * rate)
            item.update(line_id=lid, quantity=qty, unit_price=price, tax_rate=rate,
                        net=net, tax=tax, gross=net + tax, status='calculated')
        except ValueError as err:
            held_invoices.add(iid)
            issue('Lines', row, 'invalid_line', str(err))
        line_rows.append(item)
    groups = defaultdict(list)
    for line in line_rows:
        groups[line['invoice_id']].append(line)
    for iid, invoice in invoices.items():
        if iid in held_invoices or not groups[iid]:
            issue('Invoices', invoice, 'incomplete_invoice', 'Invoice held: no lines or at least one invalid line')
            continue
        group = groups[iid]
        item = dict(invoice)
        for key in ('net', 'tax', 'gross'):
            item[key] = sum((r[key] for r in group), ZERO)
        item['difference'] = item['declared_total'] - item['gross']
        item['total_check'] = 'MATCH' if item['difference'] == ZERO else 'REVIEW'
        if item['difference']:
            issue('Invoices', invoice, 'declared_mismatch', f"Declared minus calculated: {money(item['difference'])}")
        invoice_rows.append(item)
    eligible = {r['invoice_id'] for r in invoice_rows}
    payment_counts = Counter(str(r['payment_id']).strip() for r in raw['Payments'])
    for row in raw['Payments']:
        item = {'payment_id': str(row['payment_id'] or ''), 'invoice_id': str(row['invoice_id'] or '').strip(),
                'amount': '', 'source_row': row['source_row'], 'status': 'invalid'}
        try:
            amount = decimal(row['amount'], cents=True)
            item['amount'] = amount
            pid = identifier(row['payment_id'])
            iid = identifier(row['invoice_id'])
            if row['_formula']:
                raise ValueError('Input formulas are not accepted')
            item.update(payment_id=pid, invoice_id=iid)
            if payment_counts[pid] > 1:
                item['status'] = 'duplicate_id'
                issue('Payments', row, 'duplicate_payment_id', 'Every occurrence held; confirm identity before allocating')
            elif iid not in eligible:
                item['status'] = 'unallocated'
                issue('Payments', row, 'unallocated_payment', 'No complete, unique invoice to allocate against')
            else:
                item['status'] = 'allocated'
        except ValueError as err:
            issue('Payments', row, 'invalid_payment', str(err))
        payment_rows.append(item)
    for row in invoice_rows:
        row['allocated'] = sum((p['amount'] for p in payment_rows if p['status'] == 'allocated' and p['invoice_id'] == row['invoice_id']), ZERO)
        row['balance'] = row['gross'] - row['allocated']
        row['balance_status'] = 'SETTLED' if row['balance'] == ZERO else ('DUE' if row['balance'] > ZERO else 'CREDIT')
    totals = {k: money(sum((r[k] for r in invoice_rows), ZERO)) for k in ('net', 'tax', 'gross', 'declared_total', 'difference', 'allocated', 'balance')}
    categories = {s: [p for p in payment_rows if p['status'] == s] for s in ('allocated', 'duplicate_id', 'unallocated', 'invalid')}
    amounts = {s: money(sum((p['amount'] for p in rows if isinstance(p['amount'], Decimal)), ZERO)) for s, rows in categories.items()}
    raw_money = sum((p['amount'] for p in payment_rows if isinstance(p['amount'], Decimal)), ZERO)
    checks = {
        'all_line_rows_accounted_for': len(line_rows) == len(raw['Lines']),
        'all_invoice_rows_accounted_for': len(invoice_rows) + sum(i['reason'] in ('invalid_invoice', 'incomplete_invoice') for i in issues) == len(raw['Invoices']),
        'all_payment_rows_accounted_for': sum(map(len, categories.values())) == len(raw['Payments']),
        'payment_money_conserved': raw_money == sum((Decimal(v) for v in amounts.values()), ZERO),
        'net_plus_tax_equals_gross': Decimal(totals['net']) + Decimal(totals['tax']) == Decimal(totals['gross']),
        'gross_less_allocated_equals_balance': Decimal(totals['gross']) - Decimal(totals['allocated']) == Decimal(totals['balance']),
    }
    report = {'currency': 'GBP', 'rounding': 'Line net half-up to cents, then line tax half-up to cents; sum rounded lines.',
              'input_records': {k: len(v) for k, v in raw.items()},
              'output_records': {'invoice_ledger': len(invoice_rows), 'line_audit': len(line_rows), 'payment_allocations': len(payment_rows), 'exception_list': len(issues)},
              'exception_reasons': dict(Counter(i['reason'] for i in issues)),
              'checked_invoices': len(invoice_rows), 'held_invoice_rows': len(raw['Invoices']) - len(invoice_rows),
              'calculated_line_rows': sum(r['status'] == 'calculated' for r in line_rows),
              'held_line_rows': sum(r['status'] == 'held' for r in line_rows),
              'declared_mismatches': sum(r['difference'] != ZERO for r in invoice_rows),
              'credit_notes': sum(r['gross'] < ZERO for r in invoice_rows),
              'balances': dict(Counter(r['balance_status'] for r in invoice_rows)),
              'payment_rows': {k: len(v) for k, v in categories.items()},
              'duplicate_payment_ids': sum(v > 1 for v in payment_counts.values()),
              'payment_amounts': amounts, 'parsed_payment_total': money(raw_money),
              'unparseable_payment_rows': sum(not isinstance(r['amount'], Decimal) for r in payment_rows),
              'totals': totals, 'exception_rows': len(issues), 'checks': checks}
    return {'report': report, 'invoices': invoice_rows, 'lines': line_rows, 'payments': payment_rows, 'exceptions': issues}


def serial(value):
    if isinstance(value, Decimal):
        return format(value, 'f')
    raise TypeError(type(value).__name__)


def save_json(path, value):
    Path(path).write_text(json.dumps(value, indent=2, default=serial) + '\n')


def csv_file(path, rows, columns):
    with Path(path).open('w', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            # Spreadsheet applications must treat identifiers as text, including CSV exports.
            writer.writerow({k: ("'" + v if isinstance((v := row.get(k, '')), str) and v.startswith(('=', '+', '-', '@')) else v) for k in columns})


def new_book():
    book = Workbook()
    book.remove(book.active)
    book.properties.creator = 'Genco Doganay'
    book.properties.lastModifiedBy = 'Genco Doganay'
    return book


def put(sheet, row, column, value, formula=False):
    cell = sheet.cell(row, column, float(value) if isinstance(value, Decimal) else value)
    if isinstance(value, str) and not formula:
        cell.data_type = 's'
    return cell


def style(sheet, widths, header=1):
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = f'A{header + 1}'
    for i, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(i)].width = width
    for row in sheet:
        sheet.row_dimensions[row[0].row].height = 24
        for cell in row:
            cell.font = Font(name='Carlito', size=12, color='202C36')
            cell.alignment = Alignment(vertical='center')
            if cell.row == header:
                cell.fill = PatternFill('solid', fgColor='183D43')
                cell.font = Font(name='Carlito', size=12, color='FFFFFF', bold=True)
            elif cell.row % 2:
                cell.fill = PatternFill('solid', fgColor='EDF4F3')
            if isinstance(cell.value, (float, Decimal)):
                cell.number_format = '#,##0.00;[Red](#,##0.00)'
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = 'landscape'
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.print_options.horizontalCentered = True
    sheet.print_area = sheet.dimensions
    sheet.sheet_properties.outlinePr.summaryRight = False


def table(book, name, rows, columns):
    sheet = book.create_sheet(name)
    sheet.append(columns)
    for r, data in enumerate(rows, 2):
        for c, key in enumerate(columns, 1):
            put(sheet, r, c, data.get(key, ''))
    style(sheet, [max(15, min(34, len(k) + 3)) for k in columns])
    sheet.auto_filter.ref = sheet.dimensions
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = '1:1'
    return sheet


def write_outputs(result, out):
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    invoices, lines, payments = (result[k] for k in ('invoices', 'lines', 'payments'))
    cols = {
        'invoices': ['invoice_id', 'net', 'tax', 'gross', 'declared_total', 'difference', 'total_check', 'allocated', 'balance', 'balance_status', 'source_row'],
        'lines': ['invoice_id', 'line_id', 'quantity', 'unit_price', 'tax_rate', 'net', 'tax', 'gross', 'source_row', 'status'],
        'payments': ['payment_id', 'invoice_id', 'amount', 'status', 'source_row'],
        'exceptions': ['sheet', 'source_row', 'invoice_id', 'record_id', 'reason', 'detail'],
    }
    for key, columns in cols.items():
        csv_file(out / f'{key}.csv', result[key], columns)
    save_json(out / 'report.json', result['report'])
    book = new_book()
    summary = book.create_sheet('Summary')
    r = result['report']
    summary.append(['INVOICE CHECK', 'RESULT'])
    summary.append(['Invoices checked', r['checked_invoices']])
    summary.append(['Totals to review', r['declared_mismatches']])
    summary.append(['Payments held / unmatched', sum(r['payment_rows'][k] for k in ('duplicate_id', 'unallocated', 'invalid'))])
    summary.append(['Calculated gross GBP', Decimal(r['totals']['gross'])])
    summary.append(['Allocated payments GBP', Decimal(r['totals']['allocated'])])
    summary.append(['Due less credits GBP', Decimal(r['totals']['balance'])])
    style(summary, [36, 24])
    for row in summary:
        for cell in row:
            cell.font = Font(name='Carlito', size=17, bold=True, color='FFFFFF' if cell.row == 1 else '183D43')
    for i in range(1, 8):
        summary.row_dimensions[i].height = 35
    owner = book.create_sheet('Payment list')
    owner.append(['Invoice', 'Payment check', 'Due / credit', 'To review'])
    selected = []
    for state in ('SETTLED', 'DUE', 'CREDIT'):
        match = next((item for item in invoices if item['balance_status'] == state and item['gross'] > ZERO), None)
        if match is not None:
            selected.append(match)
    ordered = selected + [item for item in invoices if item not in selected]
    disputed = {item['invoice_id'] for item in payments if item['status'] in ('duplicate_id', 'invalid')}
    for n, item in enumerate(ordered, 2):
        balance = item['balance']
        status = ('Paid' if item['gross'] >= ZERO else 'Refunded') if balance == ZERO else ('Part paid' if item['allocated'] > ZERO else 'Amount due') if balance > ZERO else ('Overpaid' if item['gross'] > ZERO else 'Credit due')
        review = []
        if item['invoice_id'] in disputed:
            status = 'Check payment'
            review.append('Repeated or invalid payment entry; check before chasing payment')
        if item['difference']:
            review.append('Invoice total needs checking')
        amount = 'Nothing owed' if balance == ZERO else f'£{money(balance)} due' if balance > ZERO else f'£{money(-balance)} credit'
        for c, value in enumerate([item['invoice_id'], status, amount, '; '.join(review)], 1):
            put(owner, n, c, value)
    style(owner, [17, 19, 24, 66])
    owner.auto_filter.ref = owner.dimensions
    owner.page_setup.fitToHeight = 0
    owner.print_title_rows = '1:1'
    for row in owner:
        owner.row_dimensions[row[0].row].height = 26
        for cell in row:
            cell.font = Font(name='Carlito', size=14, bold=True, color='FFFFFF' if cell.row == 1 else '183D43')
    ledger = table(book, 'Ledger', invoices, cols['invoices'])
    maths = table(book, 'Line maths', lines, cols['lines'])
    alloc = table(book, 'Payments', payments, cols['payments'])
    exceptions_sheet = table(book, 'Exceptions', result['exceptions'], cols['exceptions'])
    exceptions_sheet.column_dimensions['F'].width = 60
    for cells in exceptions_sheet:
        exceptions_sheet.row_dimensions[cells[0].row].height = 42
        for cell in cells:
            cell.alignment = Alignment(vertical='center', wrap_text=True)
    for i, row in enumerate(lines, 2):
        if row['status'] == 'calculated':
            for c, formula in [(6, f'=ROUND(C{i}*D{i},2)'), (7, f'=ROUND(F{i}*E{i},2)'), (8, f'=ROUND(F{i}+G{i},2)')]:
                put(maths, i, c, formula, True).number_format = '#,##0.00;[Red](#,##0.00)'
            maths.cell(i, 4).number_format = '0.0000'
            maths.cell(i, 5).number_format = '0.00%'
    for i, row in enumerate(invoices, 2):
        linked_lines = [j for j, item in enumerate(lines, 2) if item['invoice_id'] == row['invoice_id']]
        for c, col in [(2, 'F'), (3, 'G'), (4, 'H')]:
            refs = ','.join(f"'Line maths'!{col}{j}" for j in linked_lines)
            put(ledger, i, c, f'=ROUND(SUM({refs}),2)', True)
        put(ledger, i, 6, f'=ROUND(E{i}-D{i},2)', True)
        put(ledger, i, 7, f'=IF(F{i}=0,"MATCH","REVIEW")', True)
        linked_payments = [j for j, item in enumerate(payments, 2) if item['invoice_id'] == row['invoice_id'] and item['status'] == 'allocated']
        refs = ','.join(f'Payments!C{j}' for j in linked_payments) or '0'
        put(ledger, i, 8, f'=ROUND(SUM({refs}),2)', True)
        put(ledger, i, 9, f'=ROUND(D{i}-H{i},2)', True)
        for c in (2, 3, 4, 5, 6, 8, 9):
            ledger.cell(i, c).number_format = '#,##0.00;[Red](#,##0.00)'
    book.save(out / 'checked-ledger.xlsx')
    views = new_book()
    fragments = []
    for idx, invoice in enumerate(invoices, 1):
        iid = invoice['invoice_id']
        sheet = views.create_sheet(f'Invoice {idx:03}')
        sheet.append([f"INVOICE {iid}", '', '', ''])
        sheet.append(['Line', 'Net GBP', 'Tax GBP', 'Gross GBP'])
        subset = [line for line in lines if line['invoice_id'] == iid]
        for line in subset:
            sheet.append([line['line_id'], float(line['net']), float(line['tax']), float(line['gross'])])
        sheet.append(['CALCULATED', float(invoice['net']), float(invoice['tax']), float(invoice['gross'])])
        sheet.append(['DECLARED', '', '', float(invoice['declared_total'])])
        sheet.append(['DIFFERENCE', '', '', float(invoice['difference'])])
        sheet.append(['PAYMENTS', '', '', float(invoice['allocated'])])
        sheet.append(['BALANCE', '', '', float(invoice['balance'])])
        style(sheet, [24, 17, 17, 17], header=2)
        for row in sheet:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = 's'
        if invoice['difference']:
            for cell in sheet[len(subset) + 5]:
                cell.fill = PatternFill('solid', fgColor='FFF0CC')
                cell.font = Font(name='Carlito', size=12, bold=True, color='804600')
        sheet['A1'].font = Font(name='Carlito', size=18, bold=True)
        sheet.merge_cells('A1:D1')
        fragments.append(f'<section><h2>{html.escape(iid)}</h2><table><tr><th>Line</th><th>Quantity</th><th>Unit price</th><th>Tax rate</th><th>Net</th><th>Tax</th><th>Gross</th></tr>' + ''.join('<tr>' + ''.join(f'<td>{html.escape(str(line[k]))}</td>' for k in ('line_id', 'quantity', 'unit_price', 'tax_rate', 'net', 'tax', 'gross')) + '</tr>' for line in subset) + '</table><p>' + ' · '.join(f'{label}: £{money(invoice[key])}' for label, key in [('Calculated', 'gross'), ('Declared', 'declared_total'), ('Difference', 'difference'), ('Payments', 'allocated'), ('Balance', 'balance')]) + f'</p><p>Source: input.xlsx, Invoices row {invoice["source_row"]}. Line sources are recorded in lines.csv.</p></section>')
    if not invoices:
        views.create_sheet('No complete invoices')['A1'] = 'No complete invoices; see exceptions.csv.'
    views.save(out / 'invoice-views.xlsx')
    (out / 'invoice-views.html').write_text('<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Invoice checks</title><style>body{font:16px system-ui;margin:24px;color:#183d43}section{margin:24px 0;padding:16px;background:#edf4f3;overflow:auto}td,th{padding:10px;text-align:right;border-bottom:1px solid #b9ccc9}th:first-child,td:first-child{text-align:left}@media print{section{break-after:page}}</style><h1>Invoice checks</h1><p>GBP. Difference = declared minus calculated. Balance = calculated gross minus allocated payments.</p>' + ''.join(fragments) + '</html>')
    text = '# Reconciliation\n\n'
    text += f"Checked invoices: {r['checked_invoices']}; line rows: {r['input_records']['Lines']}; payment rows: {r['input_records']['Payments']}.\n\n"
    text += f"Declared totals requiring review: {r['declared_mismatches']}. Credit notes: {r['credit_notes']}.\n\n"
    text += '| Invoice totals (GBP) | Amount |\n|---|---:|\n' + ''.join(f'| {key} | {value} |\n' for key, value in r['totals'].items())
    text += '\n| Payment disposition | Rows | GBP |\n|---|---:|---:|\n' + ''.join(f'| {s} | {r["payment_rows"][s]} | {r["payment_amounts"][s]} |\n' for s in r['payment_rows'])
    text += f'\nAll parsed payment rows: GBP {r["parsed_payment_total"]}. Repeated IDs are not confirmed payments: their rows stay held, including the first occurrence.\n\n'
    text += 'A negative balance means credit owed or an overpayment. Balances are calculated against line totals, while declared differences stay open for review.\n\n'
    text += '\n'.join(f'- {key}: {"PASS" if value else "FAIL"}' for key, value in r['checks'].items()) + '\n'
    (out / 'reconciliation.md').write_text(text)
    return result


def build(source, out):
    return write_outputs(reconcile(source), out)
