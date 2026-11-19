"""Run, recalculate and independently verify the reference workbooks."""
import argparse
import hashlib
import importlib.metadata
import json
import platform
import subprocess
import tempfile
import time
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
import openpyxl
from dashboard import build
from workbook_meta import scrub


def verify(path, report):
    book = openpyxl.load_workbook(path, data_only=True)
    dash = book['Dashboard']
    assert Decimal(str(dash['A3'].value)).quantize(Decimal('.01')) == Decimal(report['paid_revenue'])
    assert dash['B6'].value == report['paid_count']
    assert Decimal(str(dash['B7'].value)).quantize(Decimal('.01')) == Decimal(report['pending_amount'])
    for n, value in enumerate(report['paid_by_month'].values(), 12):
        assert Decimal(str(dash.cell(n, 2).value)).quantize(Decimal('.01')) == Decimal(value)
    start = 14 + max(1, len(report['paid_by_month']))
    for n, value in enumerate(report['paid_by_category'].values(), start):
        assert Decimal(str(dash.cell(n, 2).value)).quantize(Decimal('.01')) == Decimal(value)
    valid = list(book['ValidSales'].values)[1:]
    audit = list(book['Audit'].values)[1:]
    assert len(audit) == report['input_records']
    assert len(valid) == report['accepted_records']
    cents = sum(int(Decimal(str(r[4])) * 100) for r in valid if r[5] == 'Paid')
    assert cents == int(Decimal(report['paid_revenue']) * 100)
    assert all(report['checks'].values())
    book.close()
    formulas = openpyxl.load_workbook(path, data_only=False)
    assert len(formulas['Dashboard']._charts) == sum(bool(report[key]) for key in ('paid_by_month','paid_by_category'))
    for chart, key in zip(formulas['Dashboard']._charts, ('paid_by_month','paid_by_category')):
        ref = chart.series[0].val.numRef.f
        from openpyxl.utils.cell import range_boundaries
        _, first, _, last = range_boundaries(ref.split('!')[-1])
        assert last - first + 1 == len(report[key])
    formulas.close()


def run(input_dir, output_dir):
    start = time.perf_counter()
    out = Path(output_dir)
    report = build(input_dir, out)
    files = [('dashboard.xlsx', report)]
    if 'before' in report:
        files.append(('before-dashboard.xlsx', report['before']))
    for name, expected in files:
        with tempfile.TemporaryDirectory(prefix='.recalc-', dir=out) as tmp:
            target = Path(tmp) / name
            subprocess.run(['xlrecalc', str(out / name), str(target)], check=True, capture_output=True, text=True)
            verify(target, expected)
            (out / name).write_bytes(target.read_bytes())
            scrub(out / name)
            verify(out / name, expected)
    report['checks']['recalculated_values_match'] = True
    report['runtime_seconds'] = round(time.perf_counter() - start, 3)
    (out / 'report.json').write_text(json.dumps(report, indent=2) + '\n')
    runtime = {'executed_at_utc': datetime.now(timezone.utc).isoformat(),
               'execution': 'Python reference; LibreOffice formula recalculation with macros disabled',
               'python': platform.python_version(),
               'libreoffice': subprocess.check_output(['libreoffice', '--version'], text=True).strip(),
               'packages': {n: importlib.metadata.version(n) for n in ['openpyxl', 'XlsxWriter', 'pytest', 'Pillow']},
               'elapsed_seconds': report['runtime_seconds'],
               'input_sha256': report['input_sha256'],
               'output_sha256': {name: hashlib.sha256((out/name).read_bytes()).hexdigest() for name, _ in files},
               'excel_vba_execution': 'not_run', 'limitations': 'No desktop Excel, VBA button, or browser execution is established by this run.'}
    (out / 'runtime.json').write_text(json.dumps(runtime, indent=2) + '\n')
    lines = ['# Sales reconciliation', '', f"Paid revenue: £{report['paid_revenue']} across {report['paid_count']} paid transactions.",
             f"Pending: £{report['pending_amount']} across {report['pending_count']} transactions; excluded from paid revenue.",
             f"Cancelled: {report['cancelled_count']} rows (£{report['cancelled_amount_excluded']}) excluded from paid revenue.",
             f"Conservation: {report['input_records']} input rows = {report['accepted_records']} accepted + {report['rejected_records']} rejected.",
             '', 'Duplicate policy: exclude every occurrence of a repeated nonblank ID. Review and correct the source; no first-row guess.',
             '', '| Reason | Rows |', '|---|---:|']
    lines += [f'| {reason} | {count} |' for reason, count in report['rejection_reasons'].items()]
    lines += ['', '| Paid month | GBP |', '|---|---:|'] + [f'| {k} | {v} |' for k,v in report['paid_by_month'].items()]
    lines += ['', '| Paid category | GBP |', '|---|---:|'] + [f'| {k} | {v} |' for k,v in report['paid_by_category'].items()]
    if 'before' in report:
        lines += ['', f"Before: £{report['before']['paid_revenue']}; after: £{report['paid_revenue']}; change: £{report['paid_revenue_change']}."]
    lines += ['', '[Machine-readable figures](report.json) · [Row-by-row audit](audit.csv) · [Runtime](runtime.json)', '',
              'All totals are derived from accepted rows using decimal arithmetic, then compared with the recalculated workbook.',
              'Reference execution does not establish desktop macro execution.']
    (out / 'reconciliation.md').write_text('\n'.join(lines)+'\n')
    print(f"Paid £{report['paid_revenue']} | {report['paid_count']} paid | {report['rejected_records']} review | {report['runtime_seconds']:.3f}s")
    print('PASS: records, decimal totals and recalculated formulas agree')

if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('--input', default='input')
    p.add_argument('--output', default='output')
    a = p.parse_args()
    run(a.input, a.output)
