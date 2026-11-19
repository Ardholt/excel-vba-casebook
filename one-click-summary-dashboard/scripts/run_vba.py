"""Execute the packaged VBA refresh in Calc and compare full business outputs."""
import argparse
from datetime import datetime, timezone
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
from workbook_meta import scrub
from preserve_vba_names import preserve
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
    source_values=source.getCellRangeByPosition(0,0,4,max((r['source_row'] for r in rows),default=1)-1).getDataArray()
    for raw in rows:
        for col,key in enumerate(('sale_id','date','category','amount','status')):
            actual=source_values[raw['source_row']-1][col]
            expected=raw[key] if raw[key] is not None else ''
            assert actual==expected, ('source changed',raw['source_row'],key,actual,expected)
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


def execute(input_dir="input", output_dir="output"):
    input_dir, output_dir = Path(input_dir), Path(output_dir)
    baseline = input_dir / ("before.xlsx" if (input_dir/"before.xlsx").exists() else "input.xlsx")
    final = input_dir/"input.xlsx"
    old_rows, final_rows = read_sales(baseline), read_sales(final)
    assert final_rows[:len(old_rows)] == old_rows, "before.xlsx must be an unchanged prefix of input.xlsx"
    started=time.perf_counter()
    record={'implementation':'Packaged Forms button activated through LibreOffice XAccessibleAction; its imported actionPerformed binding calls Module1.RefreshDashboard',
            'executed_at_utc':datetime.now(timezone.utc).isoformat(),
            'libreoffice':subprocess.check_output(['libreoffice','--version'],text=True).strip(),
            'source_sha256':hashlib.sha256(Path('vba/SalesDashboard.bas').read_bytes()).hexdigest(),
            'workbook_sha256':hashlib.sha256((output_dir/'sales-dashboard.xlsm').read_bytes()).hexdigest(),
            'interaction':'LibreOffice accessibility button press, not a physical mouse click',
            'input_sha256':{p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in {baseline,final}},
            'not_exercised':['Excel desktop execution','Interactive failure dialog'],
            'excel_execution':'not_run','runs':[]}
    with tempfile.TemporaryDirectory(prefix='.vba-',dir=output_dir) as tmp:
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
            doc=desktop.loadComponentFromURL((output_dir/'sales-dashboard.xlsm').resolve().as_uri(),'_blank',0,
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
            dashboard=doc.getSheets().getByName('Dashboard')
            forms=dashboard.getDrawPage().getForms()
            assert forms.getCount()==1, 'Expected the packaged button form'
            form=forms.getByIndex(0)
            assert form.getCount()==1, 'Expected exactly one refresh button'
            model=form.getByIndex(0)
            assert model.Label=='Refresh sales'
            events=form.getScriptEvents(0)
            assert len(events)==1
            event=events[0]
            expected_binding=f'vnd.sun.star.script:{library}.Module1.RefreshDashboard?language=Basic&location=document'
            assert event.EventMethod=='actionPerformed' and event.ScriptCode==expected_binding
            control=doc.getCurrentController().getControl(model)
            control.setDesignMode(False)
            accessible=control.getAccessibleContext()
            assert accessible.getAccessibleActionCount()==1
            record['button']={'caption':model.Label, 'binding':event.ScriptCode,
                              'event':event.EventMethod, 'activation':'XAccessibleAction.doAccessibleAction(0)',
                              'installed_in_package':True}
            for phase,input_path in [('before',baseline),('after',final),('repeat',final)]:
                if phase=='after':
                    sales=doc.getSheets().getByName('Sales')
                    for raw in final_rows[len(old_rows):]:
                        for col,key in enumerate(('sale_id','date','category','amount','status')):
                            cell=sales.getCellByPosition(col,raw['source_row']-1)
                            value=raw[key]
                            if isinstance(value,(int,float)) and not isinstance(value,bool): cell.setValue(value)
                            else: cell.setString('' if value is None else str(value))
                # Every press must rebuild deliberately cleared outputs, including repeat.
                for name in ['ValidSales','Audit']:
                    doc.getSheets().getByName(name).getCellRangeByPosition(0,0,7,len(final_rows)+10).clearContents(1023)
                dashboard.getCellRangeByName('A3').setValue(-1)
                dashboard.getCellRangeByName('A10').setString('Awaiting button refresh')
                t=time.perf_counter()
                assert accessible.doAccessibleAction(0), 'Button activation refused'
                deadline=time.monotonic()+45
                while dashboard.getCellRangeByName('A10').getString()=='Awaiting button refresh':
                    if time.monotonic()>deadline: raise RuntimeError('Button did not complete refresh')
                    time.sleep(.05)
                assert dashboard.getCellRangeByName('A10').getString()=='Checked: all source rows accounted for'
                snapshot,expected=check(doc,input_path)
                snapshot.update({'phase':phase,'seconds':round(time.perf_counter()-t,3),
                                 'button_activated':True,'cleared_outputs_rebuilt':True})
                target=output_dir/f'vba-{phase}.xlsx'
                doc.storeToURL(target.resolve().as_uri(),(prop('FilterName','Calc MS Excel 2007 XML'),prop('Overwrite',True)))
                scrub(target)
                verify(target,expected)
                snapshot['saved_output']=target.name
                snapshot['saved_sha256']=hashlib.sha256(target.read_bytes()).hexdigest()
                record['runs'].append(snapshot)
                print('PASS: button',phase,snapshot['paid_revenue'])
            # Save the refreshed macro workbook through Calc, then verify its package separately.
            target=output_dir/'sales-dashboard.xlsm'
            doc.storeToURL(target.resolve().as_uri(),(prop('FilterName','Calc MS Excel 2007 VBA XML'),prop('Overwrite',True)))
            doc.close(True)
            preserve(target)
            scrub(target)
            record['vba_document_names_preserved']=['ThisWorkbook','Sheet1','Sheet2','Sheet3','Sheet4']
            record['refreshed_workbook_sha256']=hashlib.sha256(target.read_bytes()).hexdigest()
            doc=desktop.loadComponentFromURL(target.resolve().as_uri(),'_blank',0,
                 (prop('Hidden',True),prop('MacroExecutionMode',4),prop('UpdateDocMode',0)))
            dash=doc.getSheets().getByName('Dashboard')
            form=dash.getDrawPage().getForms().getByIndex(0)
            assert form.getCount()==1 and form.getByIndex(0).Label=='Refresh sales'
            assert form.getScriptEvents(0)[0].ScriptCode==expected_binding
            control=doc.getCurrentController().getControl(form.getByIndex(0))
            control.setDesignMode(False)
            dash.getCellRangeByName('A3').setValue(-1)
            dash.getCellRangeByName('A10').setString('Awaiting button refresh')
            assert control.getAccessibleContext().doAccessibleAction(0)
            deadline=time.monotonic()+45
            while dash.getCellRangeByName('A10').getString()=='Awaiting button refresh':
                if time.monotonic()>deadline: raise RuntimeError('Saved workbook button did not refresh')
                time.sleep(.05)
            check(doc, final)
            record['saved_workbook_reopened_and_button_retested']=True
            doc.close(True)
            record['result']='PASS';record['elapsed_seconds']=round(time.perf_counter()-started,3)
            (output_dir/'vba-run.json').write_text(json.dumps(record,indent=2)+'\n')
            print('PASS: actual packaged VBA, baseline, append and repeat all match the Python reference')
        except Exception as exc:
            record['result']='FAIL';record['error']=str(exc)
            (output_dir/'vba-run.json').write_text(json.dumps(record,indent=2)+'\n')
            raise
        finally:
            try: doc.close(True)
            except Exception: pass
            process.terminate()
            try: process.wait(timeout=10)
            except subprocess.TimeoutExpired: process.kill();process.wait()

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--input',default='input')
    parser.add_argument('--output',default='output')
    args=parser.parse_args()
    execute(args.input,args.output)
