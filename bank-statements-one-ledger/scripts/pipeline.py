"""Run consolidation, Calc round-trip and the business tests."""
import argparse
import hashlib
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
import pytest
from ledger import build
from verify import verify


def main():
    p=argparse.ArgumentParser(); p.add_argument('--input',default='input'); p.add_argument('--output',default='output')
    args=p.parse_args(); start=time.perf_counter(); out=Path(args.output).resolve()
    source=Path(args.input).resolve()
    if source==out or source in out.parents or out in source.parents:
        print('FAIL: input and output must be separate and non-nested.'); return 2
    try:
        report=build(args.input,out)
        print(f"Records: {report['input_records']} in, {report['output_records']} in ledger, {report['excluded_records']} rejected",flush=True)
        print(f"Balance review: {report['running_balance_mismatches']} running; {report['closing_balance_mismatches']} closing differences",flush=True)
        recalc_start=time.perf_counter()
        checked=out/'ledger-checked.xlsx'
        if shutil.which('xlrecalc'):
            result=subprocess.run(['xlrecalc',str(out/'ledger.xlsx'),str(checked)],capture_output=True,text=True)
        else:
            temp=out/'.recalc'; temp.mkdir(exist_ok=True)
            profile=out/'.lo-profile'
            result=subprocess.run(['libreoffice','-env:UserInstallation='+profile.as_uri(),'--headless','--convert-to','xlsx','--outdir',str(temp),str(out/'ledger.xlsx')],capture_output=True,text=True)
            if result.returncode==0 and (temp/'ledger.xlsx').exists(): shutil.move(temp/'ledger.xlsx',checked)
            shutil.rmtree(temp); shutil.rmtree(profile,ignore_errors=True)
        if result.returncode or not checked.exists(): raise RuntimeError('LibreOffice check failed')
        from metadata import clean_workbook
        clean_workbook(checked)
        checks=verify(out,checked)
        recalc_seconds=round(time.perf_counter()-recalc_start,6)
        test_start=time.perf_counter()
        test_temp=out/'.test-temp'
        env={**os.environ,'PYTHONDONTWRITEBYTECODE':'1','PYTHONPATH':str(Path('scripts').resolve())}
        result=subprocess.run([sys.executable,'-m','pytest','-q','-p','no:cacheprovider','--basetemp',str(test_temp),'tests'],capture_output=True,text=True,env=env)
        test_seconds=round(time.perf_counter()-test_start,6)
        match=re.search(r'(\d+) passed',result.stdout)
        count=int(match[1]) if match else 0
        shutil.rmtree(test_temp,ignore_errors=True)
        if result.returncode:
            print(result.stdout)
            raise RuntimeError('Business tests failed')
        versions={'python':platform.python_version(),'openpyxl':openpyxl.__version__,'pytest':pytest.__version__,
                  'libreoffice':subprocess.check_output(['libreoffice','--version'],text=True).strip()}
        runtime={'executed_at_utc':datetime.now(timezone.utc).isoformat(), 'environment':'Ubuntu 24.04 container',
                 'software':versions,'command':'bash scripts/inner.sh', 'input_sha256':report['input_sha256'],
                 'code_sha256':{str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted([*Path('scripts').glob('*.py'),*Path('tests').glob('*.py')])},
                 'runtime_seconds':round(time.perf_counter()-start,6),'consolidation_seconds':report['runtime_seconds'],
                 'libreoffice_check_seconds':recalc_seconds,'test_seconds':test_seconds,
                 'tests':{'passed':count,'failed':0},'checks':checks,'diagnostic_result':report['result'],
                 'result':'FAIL' if report['excluded_records'] else 'PASS',
                 'limitations':['Clean-host Docker wrapper and live GitHub/Pages checks are separate repository acceptance.',
                                'A reported balance discrepancy is a diagnostic finding, not a balanced statement.'],
                 'output_sha256':{name:hashlib.sha256((out/name).read_bytes()).hexdigest() for name in ['ledger.xlsx','ledger-checked.xlsx','input-overview.xlsx','report.json','ledger.csv','reconciliation.csv','exceptions.csv','rejected.csv','reconciliation.md']}}
        (out/'runtime.json').write_text(json.dumps(runtime,indent=2)+'\n')
        (out/'checks.txt').write_text(f"{count} business tests passed; 0 failed.\nRecords and money conserved.\nLibreOffice cell values match: {checks['nonempty_cells_compared']} nonempty cells.\nDiagnostic: {report['result']}.\n{runtime['result']}\n")
        print((out/'checks.txt').read_text(),end='')
        return 2 if report['excluded_records'] else 0
    except Exception as exc:
        detail=str(exc) if isinstance(exc,(ValueError,KeyError,RuntimeError)) else 'Unable to read inputs or complete output checks.'
        if out.exists():
            (out/'runtime.json').write_text(json.dumps({'result':'FAIL','error_type':type(exc).__name__,'detail':detail,'runtime_seconds':round(time.perf_counter()-start,6)},indent=2)+'\n')
            (out/'checks.txt').write_text('FAIL: '+detail+'\nExisting workbooks are not certified by this failed run.\n')
        print('FAIL: '+type(exc).__name__+'; '+detail)
        return 2

if __name__=='__main__': raise SystemExit(main())
