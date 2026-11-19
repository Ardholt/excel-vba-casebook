"""Build, recalculate, and verify a batch; record the execution actually performed."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
import openpyxl
from openpyxl import load_workbook
from reconcile import build, decimal, rounded, save_json


def verify_workbook(path, result):
    book = load_workbook(path, data_only=True)
    count = 0
    for index, row in enumerate(result['lines'], 2):
        if row['status'] == 'calculated':
            for col, key in ((6, 'net'), (7, 'tax'), (8, 'gross')):
                actual = book['Line maths'].cell(index, col).value
                if actual is None or rounded(decimal(actual)) != row[key]:
                    raise AssertionError(f'Line row {index}, {key}: recalculation differs')
                count += 1
    for index, row in enumerate(result['invoices'], 2):
        for col, key in ((2, 'net'), (3, 'tax'), (4, 'gross'), (6, 'difference'), (8, 'allocated'), (9, 'balance')):
            actual = book['Ledger'].cell(index, col).value
            if actual is None or rounded(decimal(actual)) != row[key]:
                raise AssertionError(f'Invoice row {index}, {key}: recalculation differs')
            count += 1
        if book['Ledger'].cell(index, 7).value != row['total_check']:
            raise AssertionError('Declared-total check differs')
        count += 1
    book.close()
    return count


def clean_properties(path):
    # Preserve formula caches: edit only package document properties.
    from zipfile import ZipFile, ZIP_DEFLATED
    from xml.etree import ElementTree as ET
    with ZipFile(path) as archive:
        payloads = {name: archive.read(name) for name in archive.namelist()}
    for name in ('docProps/core.xml', 'docProps/app.xml'):
        if name not in payloads:
            continue
        tree = ET.fromstring(payloads[name])
        for node in tree:
            tag = node.tag.split('}')[-1]
            if tag in ('creator', 'lastModifiedBy'):
                node.text = 'Genco Doganay'
            elif tag in ('Application', 'AppVersion', 'Company', 'Manager', 'Template', 'HyperlinkBase'):
                node.text = ''
        payloads[name] = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
    temporary = path.with_suffix('.clean.xlsx')
    with ZipFile(temporary, 'w', ZIP_DEFLATED) as archive:
        for name, data in payloads.items():
            archive.writestr(name, data)
    temporary.replace(path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', default='input')
    parser.add_argument('--output', default='output')
    args = parser.parse_args()
    src, out = Path(args.input).resolve(), Path(args.output).resolve()
    if src == out or src in out.parents or out in src.parents:
        raise ValueError('Input and output directories must be separate')
    start = time.perf_counter()
    source = src / 'input.xlsx'
    before_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    out.mkdir(parents=True, exist_ok=True)
    result = build(source, out)
    build_seconds = time.perf_counter() - start
    calc_start = time.perf_counter()
    with tempfile.TemporaryDirectory(prefix='.calc-', dir=out) as temp:
        target = Path(temp) / 'checked.xlsx'
        subprocess.run(['xlrecalc', str(out / 'checked-ledger.xlsx'), str(target)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        checks = verify_workbook(target, result)
        clean_properties(target)
        target.replace(out / 'checked-ledger.xlsx')
    for filename in ('invoice-views.xlsx',):
        clean_properties(out / filename)
    calc_seconds = time.perf_counter() - calc_start
    if not all(result['report']['checks'].values()):
        raise AssertionError('A conservation check did not pass')
    if hashlib.sha256(source.read_bytes()).hexdigest() != before_hash:
        raise AssertionError('Input bytes changed')
    report = result['report']
    report['checks']['libreoffice_values_match_decimal'] = True
    report['checks']['source_unchanged'] = True
    report['formula_cells_checked'] = checks
    report['runtime_seconds'] = round(time.perf_counter() - start, 3)
    save_json(out / 'report.json', report)
    with (out / 'reconciliation.md').open('a') as stream:
        stream.write(f'\nLibreOffice formula values compared: {checks}; PASS. Input bytes unchanged: PASS.\n')
    env = dict(os.environ, PYTHONDONTWRITEBYTECODE='1', CASE_RESULT=str(out), CASE_INPUT=str(src))
    test_start = time.perf_counter()
    with tempfile.TemporaryDirectory(prefix='.checks-', dir=out) as tmp:
        env['TMPDIR'] = tmp
        run = subprocess.run([sys.executable, '-m', 'pytest', '-q', '-p', 'no:cacheprovider', 'tests', '--basetemp', str(Path(tmp) / 'pytest')], env=env, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    print(run.stdout)
    if run.returncode:
        raise RuntimeError('Business checks did not pass')
    test_seconds = time.perf_counter() - test_start
    test_line = next((line.strip() for line in reversed(run.stdout.splitlines()) if ' passed' in line), '')
    lo = subprocess.run(['libreoffice', '--version'], capture_output=True, text=True, check=True).stdout.strip()
    import pytest
    runtime = {'executed_at_utc': datetime.now(timezone.utc).isoformat(),
               'execution': 'Python build and business checks; LibreOffice headless recalculation',
               'software': {'python': platform.python_version(), 'openpyxl': openpyxl.__version__, 'LibreOffice': lo, 'pytest': pytest.__version__, 'os': platform.freedesktop_os_release().get('PRETTY_NAME', 'Linux')},
               'input_sha256': before_hash, 'input_file': 'input.xlsx',
               'command': 'bash scripts/inner.sh', 'formula_cells_checked': checks,
               'checks': report['checks'], 'tests': test_line,
               'timings_seconds': {'build': round(build_seconds, 3), 'recalculate_and_compare': round(calc_seconds, 3), 'tests': round(test_seconds, 3), 'total': round(time.perf_counter() - start, 3)},
               'limitations': ['Docker entry point has not been run here; repository CI must check it.', 'Duplicate payment IDs remain held for confirmation.', 'Declared mismatches remain open; they are not silently corrected.']}
    save_json(out / 'runtime.json', runtime)
    print(f"CHECK invoices={report['checked_invoices']} declared_mismatches={report['declared_mismatches']} formula_cells={checks}")
    print('PASS: invoice arithmetic, payment allocation and conservation checks')

if __name__ == '__main__':
    main()
