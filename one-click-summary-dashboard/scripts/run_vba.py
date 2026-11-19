"""Execute the packaged VBA refresh in Calc and compare full business outputs."""
import hashlib
import json
from decimal import Decimal
from pathlib import Path
import socket
import subprocess
import tempfile
import time
import uno
from com.sun.star.beans import PropertyValue
from dashboard import analyze, read_sales
from generate_input import ADDED
from workbook_meta import scrub
from run_reference import verify


def prop(name,value):
    p=PropertyValue(); p.Name=name; p.Value=value
    return p


def norm(source):
    return '\n'.join(s.strip() for s in source.splitlines() if s.strip() and not s.strip().startswith(('Attribute ', 'Rem Attribute ')) and s.strip()!='Option VBASupport 1')


def money(value):
    return Decimal(str(value)).quantize(Decimal('.01'))


def check(doc, input_path):
    rows=read_sales(input_path); accepted,audit,report=analyze(rows)
    sheets=doc.getSheets(); d=sheets.getByName('Dashboard')
    doc.calculateAll()
    source=sheets.getByName('Sales')
    for raw in rows:
        for col,key in enumerate(('sale_id','date','category','amount','status')):
            assert source.getCellByPosition(col,raw['source_row']-1).getString()==str(raw[key] or ''), ('source changed',raw['source_row'],key)
    if money(d.getCellRangeByName('A3').getValue())!=Decimal(report['paid_revenue']):
        from collections import Counter
        actual=sheets.getByName('Audit').getCellRangeByPosition(0,1,7,len(audit)).getDataArray()
        print('Actual paid total/count:',d.getCellRangeByName('A3').getValue(),d.getCellRangeByName('B6').getValue())
        print('Audit dispositions:',dict(Counter((r[6],r[7]) for r in actual)))
        print('First audit rows:',actual[:3])
        print('First accepted rows:',sheets.getByName('ValidSales').getCellRangeByPosition(0,1,6,6).getDataArray())
    assert money(d.getCellRangeByName('A3').getValue())==Decimal(report['paid_revenue']), 'paid total'
    assert int(d.getCellRangeByName('B6').getValue())==report['paid_count'], 'paid count'
    assert money(d.getCellRangeByName('B7').getValue())==Decimal(report['pending_amount']), 'pending amount'
    assert int(d.getCellRangeByName('B8').getValue())==report['cancelled_count'], 'cancelled count'
    assert int(d.getCellRangeByName('B9').getValue())==report['rejected_records'], 'review count'
    for row,(key,amount) in enumerate(report['paid_by_month'].items(),11):
        assert d.getCellByPosition(0,row).getString()==key
        assert money(d.getCellByPosition(1,row).getValue())==Decimal(amount)
    start=13+max(1,len(report['paid_by_month']))
    for row,(key,amount) in enumerate(report['paid_by_category'].items(),start):
        assert d.getCellByPosition(0,row).getString()==key
        assert money(d.getCellByPosition(1,row).getValue())==Decimal(amount)
    v=sheets.getByName('ValidSales')
    values=v.getCellRangeByPosition(0,1,6,len(accepted)).getDataArray()
    for cells,r in zip(values,accepted):
        assert int(cells[0])==r['source_row']
        assert list(cells[1:4])==[r['sale_id'],r['date'],r['category']]
        assert money(cells[4])==r['amount']
        assert list(cells[5:])==[r['status'],r['month']]
    assert v.getCellByPosition(1,len(accepted)+1).getString()==''
    a=sheets.getByName('Audit')
    values=a.getCellRangeByPosition(0,1,7,len(audit)).getDataArray()
    for cells,r in zip(values,audit):
        assert int(cells[0])==r['source_row']
        assert list(cells[1:])==[r[k] for k in ('sale_id','date','category','amount','status','disposition','reason')], ('audit',r['source_row'],cells)
    assert a.getCellByPosition(0,len(audit)+1).getString()==''
    charts=d.getCharts()
    assert len(charts.getElementNames())==2, 'chart count'
    ranges=[]
    for name in charts.getElementNames():
        spans=charts.getByName(name).getRanges()
        assert len(spans)==1
        addr=spans[0]
        ranges.append({'start_row':addr.StartRow+1,'end_row':addr.EndRow+1,'start_column':addr.StartColumn+1,'end_column':addr.EndColumn+1})
    assert sorted((r['start_row'],r['end_row']) for r in ranges)==[(11,11+len(report['paid_by_month'])),(start,start+len(report['paid_by_category']))], ranges
    return {'paid_revenue':report['paid_revenue'],'paid_count':report['paid_count'],'pending_amount':report['pending_amount'],
            'input_records':len(audit),'accepted_records':len(accepted),'rejected_records':report['rejected_records'],
            'paid_by_month':report['paid_by_month'],'paid_by_category':report['paid_by_category'],
            'source_rows_unchanged':True,'all_accepted_rows_match':True,'all_audit_rows_match':True,'chart_ranges':ranges},report


def execute():
    started=time.perf_counter()
    record={'implementation':'VBA from output/sales-dashboard.xlsm; RefreshDashboardHeadless invokes the same RefreshSalesCore as the button',
            'libreoffice':subprocess.check_output(['libreoffice','--version'],text=True).strip(),
            'source_sha256':hashlib.sha256(Path('vba/SalesDashboard.bas').read_bytes()).hexdigest(),
            'workbook_sha256':hashlib.sha256(Path('output/sales-dashboard.xlsm').read_bytes()).hexdigest(),
            'not_exercised':['Excel desktop execution','Shape OnAction assignment and button click','Interactive failure dialog'],
            'excel_execution':'not_run','runs':[]}
    with tempfile.TemporaryDirectory(prefix='.vba-',dir='output') as tmp:
        with socket.socket() as sock:
            sock.bind(('127.0.0.1',0)); port=sock.getsockname()[1]
        process=subprocess.Popen(['libreoffice','--headless','--norestore','--nodefault','--nofirststartwizard',
            '-env:UserInstallation='+Path(tmp).resolve().as_uri(),
            f'--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        try:
            local=uno.getComponentContext()
            resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
            deadline=time.monotonic()+30
            while True:
                try:
                    ctx=resolver.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext'); break
                except Exception:
                    if time.monotonic()>=deadline: raise
                    time.sleep(.1)
            provider=ctx.ServiceManager.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider',ctx)
            config=provider.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess',
                   (prop('nodepath','/org.openoffice.Office.Common/Filter/Microsoft'),))
            enabled=[]
            def configure(node,path=''):
                for name in node.getElementNames():
                    value=node.getByName(name); full=path+'/'+name
                    if hasattr(value,'getElementNames'):
                        configure(value,full)
                    elif isinstance(value,bool) and ('excelbasic' in full.lower() or ('vba' in full.lower() and ('load' in full.lower() or 'executable' in full.lower()))):
                        node.setPropertyValue(name,True); enabled.append(full)
            configure(config); config.commitChanges()
            print('Executable VBA import settings:',enabled)
            desktop=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
            doc=desktop.loadComponentFromURL(Path('output/sales-dashboard.xlsm').resolve().as_uri(),'_blank',0,
                 (prop('Hidden',True),prop('MacroExecutionMode',4),prop('UpdateDocMode',0)))
            assert doc is not None
            libraries=doc.BasicLibraries
            print('Imported libraries:',list(libraries.getElementNames()))
            library=next(name for name in libraries.getElementNames() if name!='Standard')
            libraries.loadLibrary(library)
            imported=libraries.getByName(library).getByName('Module1')
            if norm(imported)!=norm(Path('vba/SalesDashboard.bas').read_text()):
                import difflib
                print('\n'.join(list(difflib.unified_diff(norm(Path('vba/SalesDashboard.bas').read_text()).splitlines(),norm(imported).splitlines()))[:45]))
            assert norm(imported)==norm(Path('vba/SalesDashboard.bas').read_text()), 'imported module differs'
            record['imported_source_matches']=True
            macro=doc.getScriptProvider().getScript(f'vnd.sun.star.script:{library}.Module1.RefreshDashboardHeadless?language=Basic&location=document')
            # Force real rebuilding: empty the derived rows and invalidate the displayed total.
            for name in ['ValidSales','Audit']:
                doc.getSheets().getByName(name).getCellRangeByPosition(0,0,7,1400).clearContents(1023)
            doc.getSheets().getByName('Dashboard').getCellRangeByName('A3').setValue(-1)
            for phase,input_path in [('before','input/before.xlsx'),('after','input/input.xlsx'),('repeat','input/input.xlsx')]:
                if phase=='after':
                    sales=doc.getSheets().getByName('Sales')
                    start=len(read_sales('input/before.xlsx'))+1
                    for row,values in enumerate(ADDED,start):
                        for col,value in enumerate(values): sales.getCellByPosition(col,row).setString(str(value))
                t=time.perf_counter(); result=macro.invoke((),(),())[0]
                print('VBA',phase,repr(result))
                if result!='PASS':
                    print('Sheet names:',list(doc.getSheets().getElementNames()))
                    for name in ['ValidSales','Audit','Dashboard']:
                        if doc.getSheets().hasByName(name):
                            print(name,doc.getSheets().getByName(name).getCellRangeByPosition(0,0,6,2).getDataArray())
                assert result=='PASS', str(result)
                snapshot,expected=check(doc,input_path)
                snapshot.update({'phase':phase,'seconds':round(time.perf_counter()-t,3)})
                target=Path('output')/f'vba-{phase}.xlsx'
                doc.storeToURL(target.resolve().as_uri(),(prop('FilterName','Calc MS Excel 2007 XML'),prop('Overwrite',True)))
                scrub(target)
                verify(target,expected)
                snapshot['saved_output']=str(target)
                snapshot['saved_sha256']=hashlib.sha256(target.read_bytes()).hexdigest()
                record['runs'].append(snapshot)
            doc.close(True)
            record['result']='PASS';record['elapsed_seconds']=round(time.perf_counter()-started,3)
            Path('evidence/vba-run.json').write_text(json.dumps(record,indent=2)+'\n')
            print('PASS: actual packaged VBA, baseline, append and repeat all match the Python reference')
        except Exception as exc:
            record['result']='FAIL';record['error']=str(exc)
            Path('evidence/vba-run.json').write_text(json.dumps(record,indent=2)+'\n')
            raise
        finally:
            process.terminate()
            try: process.wait(timeout=10)
            except subprocess.TimeoutExpired: process.kill();process.wait()

if __name__=='__main__': execute()
