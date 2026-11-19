"""Rebuild the included statement files deterministically."""
import csv
import json
from pathlib import Path
from datetime import date, timedelta
from decimal import Decimal
from ledger import new_book, style_sheet

base = Path('input')
base.mkdir(exist_ok=True)
a = [
 ['A001','01/10/2026','Customer receipt','', '450.00','1450.00'],
 ['A002','01/10/2026','Supplies','125.50','','1324.50'],
 ['A003','02/10/2026','Monthly charge','10.00','','1314.50'],
 ['A004','03/10/2026','Customer receipt','','220.00','1534.50'],
 ['A005','04/10/2026','Equipment','210.00','','1324.50'],
 ['A006','04/10/2026','Rounding charge','0.10','','1324.40'],
]
b = [
 ['B001','10/01/2026','Customer receipt', '450.00','950.00'],
 ['B002','10/01/2026','Rent', '-500.00','450.00'],
 ['B003','10/02/2026','Monthly charge', '-10.00','440.00'],
 ['B004','10/03/2026','Supplier refund','25.50','465.50'],
 ['B005','10/04/2026','Supplies','-90.00','375.50'],
 ['B006','10/04/2026','Rounding charge','-0.10','375.40'],
]
c = [
 ['C001','2026-10-01','Customer receipt','100.00','IN','850.00'],
 ['C002','2026-10-01','Supplies','200.00','OUT','650.00'],
 ['C003','2026-10-02','Monthly charge','15.00','OUT','635.00'],
 ['C004','2026-10-03','Customer receipt','70.00','IN','705.00'],
 ['C005','2026-10-04','Equipment','200.00','OUT','505.00'],
 ['C006','2026-10-04','No-charge adjustment','0.00','OUT','504.75'],
]
# Continue the opening transactions through the whole quarter.
# Keep each account independent, including equal receipts on separate accounts.
c[-1][-1] = '505.00'
all_rows = [('layout-a.csv', a, 'A', 132440), ('layout-b.xlsx', b, 'B', 37540), ('layout-c.csv', c, 'C', 50500)]
closing = {}
for account_index, (filename, rows, prefix, balance) in enumerate(all_rows):
    for day_index in range(88):
        day = date(2026, 10, 5) + timedelta(days=day_index)
        receipts = 18500 + ((day_index * 37 + account_index * 53) % 17000)
        supplies = -(6200 + ((day_index * 71 + account_index * 101) % 8700))
        charge = -10 if day.day != 1 else -1500
        movements = [('Customer receipt', receipts), ('Supplies', supplies), ('Bank charge', charge)]
        if prefix == 'C' and day == date(2026, 12, 31):
            movements[-1] = ('No-charge adjustment', 0)
        for description, amount in movements:
            balance += amount
            identity = f'{prefix}{len(rows)+1:03d}'
            value = format(Decimal(amount) / 100, '.2f')
            reported = format(Decimal(balance) / 100, '.2f')
            if prefix == 'A':
                rows.append([identity, day.strftime('%d/%m/%Y'), description,
                             format(Decimal(-amount)/100,'.2f') if amount<0 else '',
                             value if amount>0 else '', reported])
            elif prefix == 'B':
                rows.append([identity, day.strftime('%m/%d/%Y'), description, value, reported])
            else:
                rows.append([identity, day.isoformat(), description,
                             format(Decimal(abs(amount))/100,'.2f'), 'IN' if amount>0 else 'OUT', reported])
    # Only the final reported and declared balance of C is wrong. No money is inserted.
    if prefix == 'C':
        rows[-1][-1] = format(Decimal(balance-25)/100,'.2f')
    closing[filename] = rows[-1][-1]

from ledger import SCHEMAS
for name, rows in [('layout-a.csv', a), ('layout-c.csv', c)]:
    with (base / name).open('w', newline='') as f:
        w=csv.writer(f); w.writerow(SCHEMAS[name]); w.writerows(rows)
book=new_book(); ws=book.create_sheet('Transactions'); ws.append(SCHEMAS['layout-b.xlsx'])
for row in b: ws.append(row)
style_sheet(ws); book.save(base/'layout-b.xlsx')
from metadata import clean_workbook
clean_workbook(base/'layout-b.xlsx')
settings={name: {'account': code, 'currency': 'GBP', 'opening_balance': op, 'closing_balance': closing[name], 'date_order': order}
          for name,code,op,order in [('layout-a.csv','ACCT-A','1000.00','DMY'),('layout-b.xlsx','ACCT-B','500.00','MDY'),('layout-c.csv','ACCT-C','750.00','DMY')]}
(base/'accounts.json').write_text(json.dumps(settings,indent=2)+'\n')
