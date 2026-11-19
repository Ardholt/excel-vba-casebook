import json
import os
import time
from pathlib import Path
import pytest
_started = time.perf_counter()
_results = []


def pytest_runtest_logreport(report):
    if report.when == 'call' or (report.when == 'setup' and report.failed):
        _results.append({'test':report.nodeid.split('::',1)[-1], 'outcome':report.outcome})


def pytest_sessionfinish(session, exitstatus):
    target = os.environ.get('CASE_TEST_REPORT')
    if target:
        Path(target).write_text(json.dumps({'status':'PASS' if exitstatus==0 else 'FAIL','pytest':pytest.__version__, 'tests':len(_results), 'passed':sum(r['outcome']=='passed' for r in _results), 'duration_seconds':round(time.perf_counter()-_started,4), 'results':_results},indent=2)+'\n')
