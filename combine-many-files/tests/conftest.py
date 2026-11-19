import json
from pathlib import Path
import time

_started = time.perf_counter()


def pytest_sessionfinish(session, exitstatus):
    reporter = session.config.pluginmanager.get_plugin('terminalreporter')
    counts = {key: len(reporter.stats.get(key, [])) for key in ('passed','failed','error','skipped')}
    result = {'tests_collected': session.testscollected, **counts,
              'exit_code': int(exitstatus), 'seconds': round(time.perf_counter()-_started, 3),
              'command': 'python3 -m pytest -q tests', 'environment': 'Linux container'}
    dest = Path('evidence')
    dest.mkdir(exist_ok=True)
    (dest/'tests.json').write_text(json.dumps(result, indent=2)+'\n')
