"""Open the packaged candidate with macros disabled and inspect its workbook data."""
import json
from pathlib import Path
import socket
import subprocess
import tempfile
import time
import uno
from com.sun.star.beans import PropertyValue


def prop(name,value):
    p=PropertyValue(); p.Name=name; p.Value=value
    return p

expected=json.loads(Path('output/report.json').read_text())['before']
start=time.perf_counter()
with tempfile.TemporaryDirectory(prefix='.macro-open-',dir='output') as tmp:
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
                ctx=resolver.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext')
                break
            except Exception:
                if time.monotonic()>=deadline: raise
                time.sleep(.1)
        desktop=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
        doc=desktop.loadComponentFromURL(Path('output/sales-dashboard.xlsm').resolve().as_uri(),'_blank',0,
             (prop('Hidden',True),prop('ReadOnly',True),prop('MacroExecutionMode',0),prop('UpdateDocMode',0)))
        assert doc is not None
        doc.calculateAll()
        sheets=doc.getSheets()
        assert list(sheets.getElementNames()) == ['Dashboard','Sales','ValidSales','Audit']
        dashboard=sheets.getByName('Dashboard')
        assert round(dashboard.getCellRangeByName('A3').getValue(),2)==float(expected['paid_revenue'])
        assert dashboard.getCellRangeByName('B6').getValue()==expected['paid_count']
        doc.close(True)
        record=json.loads(Path('evidence/macro-package.json').read_text())
        record['opened_in']=subprocess.check_output(['libreoffice','--version'],text=True).strip()
        record['opened_with_macros_disabled']=True
        record['baseline_revenue_checked']=expected['paid_revenue']
        record['baseline_paid_count_checked']=expected['paid_count']
        record['open_check_seconds']=round(time.perf_counter()-start,3)
        Path('evidence/macro-package.json').write_text(json.dumps(record,indent=2)+'\n')
        print('PASS: candidate XLSM opens; baseline revenue and transaction count agree')
    finally:
        process.terminate()
        try: process.wait(timeout=10)
        except subprocess.TimeoutExpired: process.kill(); process.wait()
