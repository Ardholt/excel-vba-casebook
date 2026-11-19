"""Execute the shipped module in Calc; export its cells through UNO, not VBA SaveAs."""
import json
import os
import platform
from pathlib import Path
import shutil
import subprocess
import tempfile
import time
import hashlib
import socket
import uno
from com.sun.star.beans import PropertyValue
import openpyxl
from combine import collect
from metadata import scrub


def props(**kwargs):
    return tuple(PropertyValue(Name=k, Value=v) for k, v in kwargs.items())


def normal(value):
    return None if value == '' else value


def values(doc, sheet):
    sh = doc.Sheets.getByName(sheet)
    cursor = sh.createCursor()
    cursor.gotoEndOfUsedArea(False)
    area = cursor.RangeAddress
    data = [list(map(normal, row)) for row in sh.getCellRangeByPosition(0, 0, area.EndColumn, area.EndRow).getDataArray()]
    while data and all(v is None for v in data[-1]):
        data.pop()
    if data:
        width = max((i + 1 for row in data for i, v in enumerate(row) if v is not None), default=0)
        data = [row[:width] for row in data]
    return data


def compare(doc, folder):
    headers, rows, manifest = collect(folder)
    got = values(doc, 'Combined')
    expected = [headers] + [[r.get(h) for h in headers] for r in rows]
    mismatches = [(i + 1, a, b) for i, (a, b) in enumerate(zip(got, expected)) if a != b]
    assert got == expected, f'Combined cell mismatch; received {len(got)-1} rows; first differences: {mismatches[:3]}'
    actual_manifest = values(doc, 'Manifest')
    expected_manifest = [['file', 'status', 'rows', 'reason']] + [[normal(v) for v in r] for r in manifest]
    assert actual_manifest == expected_manifest, 'Manifest mismatch'
    return len(rows), hashlib.sha256(json.dumps(got, sort_keys=True).encode()).hexdigest()


started = time.perf_counter()
Path('evidence').mkdir(exist_ok=True)
Path('output').mkdir(exist_ok=True)
source = Path('vba/CombineFiles.bas').read_text()
with tempfile.TemporaryDirectory(prefix='combine-') as td:
    tmp = Path(td)
    sock = socket.socket()
    sock.bind(('127.0.0.1', 0))
    port = sock.getsockname()[1]
    sock.close()
    process = subprocess.Popen(['libreoffice', '-env:UserInstallation=' + (tmp/'profile').as_uri(), '--headless', '--norestore', '--nodefault', '--nofirststartwizard', f'--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    doc = None
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        for _ in range(100):
            try:
                context = resolver.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext')
                break
            except Exception:
                time.sleep(.1)
        else:
            raise RuntimeError('Calc listener did not start')
        desktop = context.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', context)
        doc = desktop.loadComponentFromURL(Path('combine.xlsm').resolve().as_uri(), '_blank', 0, props(Hidden=True, MacroExecutionMode=4, UpdateDocMode=0))
        library_name = next(n for n in doc.BasicLibraries.getElementNames() if n == 'VBAProject')
        doc.BasicLibraries.loadLibrary(library_name)
        imported = doc.BasicLibraries.getByName(library_name).getByName('Module1')
        def source_lines(s):
            return [line.strip() for line in s.splitlines() if line.strip() and not line.strip().startswith(('Attribute ', 'Option VBASupport', 'Rem Attribute'))]
        assert source_lines(source) == source_lines(imported), 'Embedded module differs from BAS source'
        script = doc.getScriptProvider().getScript('vnd.sun.star.script:VBAProject.Module1.RefreshFolder?language=Basic&location=document')
        cfg = doc.Sheets.getByName('Config')
        source_folder = Path('input/files').resolve()
        cfg.getCellRangeByName('B2').String = str(source_folder)
        timings = []
        baseline = None
        for label in ['first_refresh', 'repeat_refresh']:
            t = time.perf_counter()
            print('Executing refresh', flush=True)
            script.invoke((), (), ())
            assert cfg.getCellRangeByName('B3').String == 'Complete', cfg.getCellRangeByName('B3').String
            count, digest = compare(doc, source_folder)
            if baseline:
                assert digest == baseline, 'Repeat changed business result'
            baseline = digest
            timings.append({'operation': label, 'seconds': round(time.perf_counter()-t, 3), 'rows': count, 'cells_equal_reference': True})
        cfg.getCellRangeByName('B2').String = 'input/files'
        doc.DocumentProperties.Author = 'Genco Doganay'
        doc.DocumentProperties.ModifiedBy = 'Genco Doganay'
        doc.storeToURL(Path('output/vba-result.xlsx').resolve().as_uri(), props(FilterName='Calc MS Excel 2007 XML', Overwrite=True))
        scrub('output/vba-result.xlsx')
        extra = tmp/'new files'
        shutil.copytree(source_folder, extra)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Data'
        ws.append(['amount', 'record_id', 'category', 'date', 'new_field'])
        ws.append([75.25, 'NEW-0001', 'Service', '2025-09-01', 'retained'])
        wb.save(extra/'51_new_branch.xlsx')
        cfg.getCellRangeByName('B2').String = str(extra)
        t = time.perf_counter()
        print('Executing added-file refresh', flush=True)
        script.invoke((), (), ())
        assert cfg.getCellRangeByName('B3').String == 'Complete'
        count, _ = compare(doc, extra)
        assert count == timings[0]['rows'] + 1
        timings.append({'operation': 'added_file_refresh', 'seconds': round(time.perf_counter()-t,3), 'rows':count, 'cells_equal_reference':True})
        result = {'engine': subprocess.check_output(['libreoffice','--version'],text=True).strip(),
                  'macro': 'Module1.RefreshFolder', 'timing_scope': 'Each refresh timing includes macro invocation and source-to-cell comparison.', 'source_sha256': hashlib.sha256(source.encode()).hexdigest(),
                  'embedded_source_matches': True, 'runs': timings, 'export': 'UNO XLSX export; VBA SaveAs was not used',
                  'excel_run': 'Not yet run in Microsoft Excel', 'power_query_run': 'Not run',
                  'seconds': round(time.perf_counter()-started,3)}
        Path('evidence/vba-run.json').write_text(json.dumps(result,indent=2)+'\n')
        runtime = {'python':platform.python_version(), 'openpyxl':openpyxl.__version__, 'libreoffice':result['engine'],
                   'vba_packager':'pyOpenVBA 6.6.0', 'reference':json.loads(Path('output/report.json').read_text())['reference_seconds'],
                   'vba_seconds':result['seconds'], 'execution':'Linux container; Python reference and actual VBA module in LibreOffice headless',
                   'microsoft_excel':'Not yet run', 'power_query':'Source provided; not run', 'vba_xlsx_export':'UNO; not VBA SaveAs'}
        Path('evidence/runtime.json').write_text(json.dumps(runtime,indent=2)+'\n')
        print(json.dumps(result,indent=2))
    finally:
        if doc:
            doc.close(True)
        process.terminate()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
