"""Rebuild the committed input batch deterministically."""
from decimal import Decimal
from pathlib import Path
from reconcile import new_book, table, rounded, style


def generate(destination):
    lines, invoices, payments = [], [], []
    for i in range(1, 121):
        iid = f'INV-{i:04}'
        gross = Decimal('0')
        for j in range(1, 4):
            quantity = Decimal(-1 if i % 12 == 0 else 1) * Decimal((i + j) % 5 + 1)
            price = Decimal(835 + ((i * 73 + j * 117) % 8000)) / 100
            if j == 1 and i % 7 == 0:
                price += Decimal('0.005')
            rate = [Decimal('0'), Decimal('0.05'), Decimal('0.20')][(i + j) % 3]
            net = rounded(quantity * price)
            gross += net + rounded(net * rate)
            lines.append(dict(invoice_id=iid, line_id=f'L{j:02}', quantity=str(quantity), unit_price=str(price), tax_rate=str(rate)))
        declared = gross + (Decimal('2.35') if i % 17 == 0 else Decimal('0'))
        invoices.append(dict(invoice_id=iid, declared_total=str(declared)))
        if i % 12 == 0:
            if i % 24 == 0:
                payments.append(dict(payment_id=f'PAY-{i:04}', invoice_id=iid, amount=str(gross)))
        elif i % 5 == 0:
            payments.append(dict(payment_id=f'PAY-{i:04}', invoice_id=iid, amount=str(rounded(gross / 2))))
        elif i % 11 == 0:
            payments.append(dict(payment_id=f'PAY-{i:04}', invoice_id=iid, amount=str(gross + Decimal('18.40'))))
        elif i % 9 == 0:
            first = rounded(gross / 3)
            payments.extend([dict(payment_id=f'PAY-{i:04}-A', invoice_id=iid, amount=str(first)), dict(payment_id=f'PAY-{i:04}-B', invoice_id=iid, amount=str(gross - first))])
        else:
            payments.append(dict(payment_id=f'PAY-{i:04}', invoice_id=iid, amount=str(gross)))
    for number in (8, 34, 61):
        source = next(p for p in payments if p['payment_id'] == f'PAY-{number:04}')
        copy = dict(source)
        if number == 34:
            copy['amount'] = str(Decimal(copy['amount']) + Decimal('9.75'))
        payments.append(copy)
    for i, amount in enumerate(('45.25', '62.10', '-15.00', '90.80'), 1):
        payments.append(dict(payment_id=f'UN-{i:03}', invoice_id=f'INV-9{i:03}', amount=amount))
    book = new_book()
    overview = book.create_sheet('Overview')
    overview.append(['INPUTS TO CHECK', 'ROWS / GBP'])
    overview.append(['Invoice headers', len(invoices)])
    overview.append(['Invoice lines', len(lines)])
    overview.append(['Payment entries', len(payments)])
    overview.append(['Declared invoice total', float(sum(Decimal(i['declared_total']) for i in invoices))])
    overview.append(['Payment entries total', float(sum(Decimal(p['amount']) for p in payments))])
    overview.append(['Totals and allocations', 'Not yet checked'])
    style(overview, [36, 24])
    from openpyxl.styles import Font
    for row in overview:
        for cell in row:
            cell.font = Font(name='Carlito', size=17, bold=True, color='FFFFFF' if cell.row == 1 else '183D43')
    for i in range(1, 8):
        overview.row_dimensions[i].height = 35
    for title, rows, headers in [('Lines', lines, ['invoice_id','line_id','quantity','unit_price','tax_rate']), ('Invoices', invoices, ['invoice_id','declared_total']), ('Payments', payments, ['payment_id','invoice_id','amount'])]:
        table(book, title, rows, headers)
    Path(destination).parent.mkdir(parents=True, exist_ok=True)
    book.save(destination)
    from run_pipeline import clean_properties
    clean_properties(Path(destination))

if __name__ == '__main__':
    generate(Path('input/input.xlsx'))
