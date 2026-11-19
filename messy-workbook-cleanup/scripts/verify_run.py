"""Run the same business checks and Calc round-trip on any accepted input."""
import argparse
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
import pytest
from normalize_records import run
from workbook_meta import scrub_properties


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def values(path):
    book = openpyxl.load_workbook(path, data_only=True)
    return {ws.title: list(ws.values) for ws in book}



def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', default='input')
    parser.add_argument('--output', default='output')
    args = parser.parse_args()
    started = time.perf_counter()
    source, out = Path(args.input).resolve(), Path(args.output).resolve()
    original_hashes = {p.name: digest(p) for p in source.iterdir() if p.is_file()}
    first = run(source, out)
    with tempfile.TemporaryDirectory(prefix='.checks-', dir=out) as temporary:
        temp = Path(temporary)
        second = run(source, temp / 'repeat')
        assert first['business_result_sha256'] == second['business_result_sha256'], 'Repeat result changed'
        clean_source = temp / 'clean-source'
        clean_source.mkdir()
        shutil.copyfile(out / 'customer-register.xlsx', clean_source / 'input.xlsx')
        shutil.copyfile(source / 'rules.json', clean_source / 'rules.json')
        third = run(clean_source, temp / 'clean-again')
        assert first['business_result_sha256'] == third['business_result_sha256'], 'Cleaned input changed'
        assert third['duplicate_records'] == third['rejected_records'] == third['field_changes'] == 0
        lo = subprocess.check_output(['libreoffice', '--version'], text=True).strip()
        lo_start = time.perf_counter()
        lo_dir = temp / 'calc'
        lo_dir.mkdir()
        process = subprocess.run(['libreoffice', '-env:UserInstallation=' + (temp / 'profile').as_uri(),
                                  '--headless', '--convert-to', 'xlsx', '--outdir', str(lo_dir), str(out / 'customer-register.xlsx')],
                                 stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=120)
        roundtrip = lo_dir / 'customer-register.xlsx'
        if process.returncode or not roundtrip.exists():
            raise RuntimeError('LibreOffice round-trip failed')
        assert values(out / 'customer-register.xlsx') == values(roundtrip), 'Calc changed workbook cell values'
        shutil.copyfile(roundtrip, out / 'calc-checked.xlsx')
        scrub_properties(out / 'calc-checked.xlsx')
        lo_seconds = time.perf_counter() - lo_start
        test_start = time.perf_counter()
        test_run = subprocess.run([sys.executable, '-m', 'pytest', '-q', '-p', 'no:cacheprovider', 'tests',
                                   '--basetemp', str(temp / 'pytest')], text=True, stdout=subprocess.PIPE,
                                  stderr=subprocess.STDOUT, env={**os.environ, 'PYTHONDONTWRITEBYTECODE': '1'})
        if test_run.returncode:
            print(test_run.stdout)
            raise RuntimeError('business tests failed')
        # Successful compact pytest output has no machine paths.
        (out / 'test-results.txt').write_text(test_run.stdout)
        print(test_run.stdout.strip())
        test_seconds = time.perf_counter() - test_start
    assert original_hashes == {p.name: digest(p) for p in source.iterdir() if p.is_file()}, 'Source files changed'
    runtime = {
        'checked_utc': datetime.now(timezone.utc).isoformat(),
        'software': {'python': platform.python_version(), 'openpyxl': openpyxl.__version__,
                     'pytest': pytest.__version__, 'libreoffice': lo},
        'operating_system': {k: platform.freedesktop_os_release().get(k) for k in ('NAME', 'VERSION_ID')},
        'execution': 'Linux container; Python cleanup and tests; LibreOffice headless XLSX load and export',
        'checks': {'repeat_business_result_unchanged': True, 'cleaned_input_no_further_changes': True,
                   'libreoffice_all_sheet_values_preserved': True, 'source_files_unchanged': True},
        'seconds': {'cleanup': first['runtime_seconds'], 'libreoffice': round(lo_seconds, 4),
                    'tests': round(test_seconds, 4), 'full_run': round(time.perf_counter() - started, 4)},
        'artifacts_sha256': {f: digest(out / f) for f in ['customer-register.xlsx', 'calc-checked.xlsx', 'report.json', 'test-results.txt']},
        'limitations': ['Docker entrypoint and hosted CI are checked separately.',
                       'Currency signs are presentation only; inputs must use one currency.',
                       'No fuzzy identity merging; rejected amounts are excluded from totals.']}
    (out / 'runtime.json').write_text(json.dumps(runtime, indent=2) + '\n')
    print('Record reconciliation: PASS')
    print('Amount reconciliation: PASS')
    print('Repeat and cleaned-input checks: PASS')
    print('LibreOffice cell-value round-trip: PASS')
    print(f"PASS — {first['input_records']} source = {first['kept_records']} kept + {first['duplicate_records']} duplicate-linked + {first['rejected_records']} rejected")


if __name__ == '__main__':
    main()
