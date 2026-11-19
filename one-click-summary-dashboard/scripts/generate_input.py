"""Rebuild the included sales inputs from fixed arithmetic."""
from pathlib import Path
import xlsxwriter
from workbook_meta import scrub

HEADERS = ['sale_id', 'date', 'category', 'amount', 'status']

def rows():
    import calendar
    import random
    rng = random.Random(20251119)
    categories = ['Hardware', 'Lighting', 'Garden', 'Paint', 'Tools', 'Household']
    monthly_counts = [78, 82, 93, 96, 103, 111, 107, 115, 109, 117, 130, 159]
    result = []
    for month, count in enumerate(monthly_counts, 1):
        days = sorted(rng.randint(1, calendar.monthrange(2025, month)[1]) for _ in range(count))
        for day in days:
            n = len(result) + 1
            category = rng.choice(categories)
            cents = rng.randint(450, 24500)
            status = rng.choices(['Paid', 'Pending', 'Cancelled'], weights=[88, 8, 4])[0]
            if status == 'Paid' and rng.random() < .045:
                cents = -cents
            result.append([f'S{n:05}', f'2025-{month:02}-{day:02}', category, f'{cents/100:.2f}', status])
    for i, position in enumerate([5, 101, 209, 307, 411, 509, 602, 711, 819, 907, 1108, 1232]):
        duplicate = list(result[position])
        if i % 3 == 0:
            duplicate[3] = '199.95'
        result.append(duplicate)
    result += [['S09001', '2025-02-30', 'Hardware', '80.00', 'Paid'],
               ['S09002', '2025-06-18', 'Tools', '84.1x', 'Paid'],
               ['S09003', '2025-06-19', 'Paint', '64.00', 'Awaiting'],
               ['S09004', '2025-06-20', '', '92.00', 'Paid'],
               ['S09005', '31/07/2025', 'Garden', '18.00', 'Paid'],
               ['S09006', '2025-08-12', 'Lighting', '19.999', 'Paid'],
               ['', '2025-09-14', 'Household', '26.00', 'Paid'],
               ['S09008', '', 'Hardware', '33.00', 'Paid']]
    return result

ADDED = [[f'S{2001+i:05}', f'2026-01-{i//3+2:02}', category, amount, status]
         for i, (category, amount, status) in enumerate([
             ('Repairs', '185.50', 'Paid'), ('Hardware', '-45.25', 'Paid'),
             ('Lighting', '64.90', 'Paid'), ('Tools', '129.95', 'Paid'),
             ('Paint', '27.49', 'Paid'), ('Household', '18.75', 'Paid'),
             ('Tools', '299.00', 'Paid'), ('Garden', '87.40', 'Paid'),
             ('Hardware', '52.80', 'Paid'), ('Repairs', '140.00', 'Paid'),
             ('Repairs', '250.00', 'Pending'), ('Lighting', '125.00', 'Cancelled')])]

def write(path, data):
    with xlsxwriter.Workbook(path, {'strings_to_formulas': False, 'strings_to_urls': False}) as book:
        book.set_properties({'author': '', 'company': '', 'comments': ''})
        sheet = book.add_worksheet('Sales')
        sheet.write_row(0, 0, HEADERS)
        for n, row in enumerate(data, 1):
            sheet.write_row(n, 0, row)
        sheet.set_column('A:C', 18)
        sheet.set_column('D:E', 16)
        sheet.freeze_panes(1, 0)
        sheet.autofilter(0, 0, len(data), 4)
    scrub(path)

if __name__ == '__main__':
    dest = Path('input')
    dest.mkdir(exist_ok=True)
    write(dest / 'before.xlsx', rows())
    write(dest / 'input.xlsx', rows() + ADDED)
    write(dest / 'new-sales.xlsx', ADDED)
