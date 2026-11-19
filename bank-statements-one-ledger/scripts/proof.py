"""Narrow workbook views for readable sheet renders."""
from openpyxl.styles import Alignment, Font, PatternFill
from ledger import NAVY, TEAL, PALE, RED, money, pennies


def setup(ws, end):
    ws.sheet_view.showGridLines=False
    ws.column_dimensions['A'].width=24
    ws.column_dimensions['B'].width=32
    for row in ws:
        ws.row_dimensions[row[0].row].height=28
        for c in row:
            c.font=Font(name='Calibri',size=15,color=NAVY)
            c.alignment=Alignment(vertical='center')
            if c.row%2==0: c.fill=PatternFill('solid',fgColor=PALE)
    ws.print_area=f'A1:B{end}'
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.paperSize=ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth=1
    ws.page_setup.fitToHeight=1
    ws.print_options.horizontalCentered=True


def band(ws,row):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2)
    ws.cell(row,1).font=Font(name='Calibri',size=16,bold=True,color='FFFFFF')
    for c in ws[row]: c.fill=PatternFill('solid',fgColor=NAVY)
    ws.row_dimensions[row].height=34


def add_sources(book, sources):
    ws=book.create_sheet('Statement layouts',0)
    bands=[]
    for filename, rownum, row in sources:
        bands.append(ws.max_row+1 if bands else 1)
        ws.append([filename+' · row '+str(rownum),''])
        for field,value in row.items(): ws.append([field,value])
        ws.append(['',''])
    if not sources:
        ws.append(['No source transactions','']); bands.append(1)
    setup(ws,max(1,ws.max_row-1))
    for r in bands: band(ws,r)
    for row in ws:
        for c in row:
            if isinstance(c.value,str): c.data_type='s'


def add_trace(book, ledger):
    ws=book.create_sheet('Trace a balance')
    row=next((r for r in ledger if r['check']=='MISMATCH'),ledger[-1] if ledger else None)
    ws.append(['FROM STATEMENT TO LEDGER',''])
    if row:
        for field,title in [('account','Account'),('source_file','Source file'),('source_row','Source row'),('id','Transaction'),('date','Date'),('amount','Signed amount'),('reported_balance','Reported balance'),('calculated_balance','Calculated balance'),('difference','Difference'),('check','Check')]:
            value=row[field]
            if field in ('amount','reported_balance','calculated_balance','difference') and value!='': value='£'+value
            ws.append([title,value])
        ws.append(['Source amount preserved.',''])
        ws.append(['No extra transaction added.',''])
    setup(ws,ws.max_row)
    band(ws,1)
    for r in (12,13):
        if r<=ws.max_row:
            ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
            ws.cell(r,1).font=Font(name='Calibri',size=13,italic=True,color=NAVY)
    if row and row['check']=='MISMATCH':
        for r in (10,11):
            ws.cell(r,2).font=Font(name='Calibri',size=15,bold=True,color=RED)


def compact_balance(book, name, title, headers, values):
    ws=book.create_sheet(name)
    ws.sheet_view.showGridLines=False
    for col in ('A','B','C'): ws.column_dimensions[col].width=25
    ws.append([title,'','']); ws.merge_cells('A1:C1')
    ws.append(headers); ws.append(values)
    for i in range(1,4):
        ws.row_dimensions[i].height=41
        for c in ws[i]:
            c.alignment=Alignment(vertical='center')
            c.font=Font(name='Calibri',size=18,bold=i==1,color='FFFFFF' if i==1 else NAVY)
            c.fill=PatternFill('solid',fgColor=NAVY if i==1 else PALE)
    ws.print_area='A1:C3'
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation='landscape'
    ws.page_setup.paperSize=ws.PAPERSIZE_A5
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1
    return ws


def add_hero_after(book, ledger):
    row=next((r for r in ledger if r['check']=='MISMATCH'), ledger[-1] if ledger else None)
    if row:
        compact_balance(book,'Closing balance proof',row['account']+' · closing balance check',
                        ['Reported','Calculated','Difference'],
                        [row['reported_balance'],row['calculated_balance'],row['difference']])


def add_hero_before(book, rows):
    if rows:
        rownum,row=rows[-1]
        compact_balance(book,'Source balance proof',row['reference']+' · statement row '+str(rownum),
                        ['Amount','Direction','Balance'],[row['amount'],row['direction'],row['balance']])
