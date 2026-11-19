"""Check Docker arguments without needing Docker inside the test container."""
import json
import os
from pathlib import Path
import shutil
import subprocess

import pytest


@pytest.fixture
def launcher(tmp_path):
    case = tmp_path / 'case with spaces'
    (case / 'input/files').mkdir(parents=True)
    (case / 'input/files/source.xlsx').write_bytes(b'untouched source')
    (case / 'output').mkdir()
    (case / 'output/published.xlsx').write_bytes(b'published result')
    shutil.copyfile(Path(__file__).resolve().parents[1] / 'run.sh', case / 'run.sh')
    bin_dir = tmp_path / 'bin'
    bin_dir.mkdir()
    docker = bin_dir / 'docker'
    docker.write_text('#!/usr/bin/env python3\nimport json, os, sys\n'
                      'with open(os.environ["DOCKER_LOG"], "a") as f:\n'
                      '    f.write(json.dumps(sys.argv[1:]) + "\\n")\n')
    docker.chmod(0o755)
    log = tmp_path / 'docker.jsonl'
    env = {**os.environ, 'PATH': str(bin_dir) + os.pathsep + os.environ['PATH'],
           'DOCKER_LOG': str(log)}
    def run(*args):
        return subprocess.run(['bash', str(case / 'run.sh'), *map(str, args)],
                              env=env, capture_output=True, text=True)
    return case, log, run


def test_default_uses_readonly_sources_and_new_results_each_time(launcher):
    case, log, run = launcher
    for _ in range(2):
        result = run()
        assert result.returncode == 0, result.stderr
    calls = [json.loads(line) for line in log.read_text().splitlines()]
    runs = [c for c in calls if c[0] == 'run']
    outputs = []
    for args in runs:
        mounts = [args[i + 1] for i, value in enumerate(args) if value == '-v']
        assert mounts[:2] == [f'{case}:/case:ro', f'{case}/input/files:/source:ro']
        assert len(mounts) == 3
        destination, target, mode = mounts[2].rsplit(':', 2)
        assert (target, mode) == ('/result', 'rw')
        assert Path(destination).parent == case / 'runs'
        assert list(Path(destination).iterdir()) == []
        assert args[-3:] == ['--demo', '/source', '/result']
        outputs.append(destination)
    assert len(set(outputs)) == 2
    assert (case / 'output/published.xlsx').read_bytes() == b'published result'
    assert (case / 'input/files/source.xlsx').read_bytes() == b'untouched source'


def test_custom_sources_readonly_and_fresh_output(launcher, tmp_path):
    case, log, run = launcher
    source, output = tmp_path / 'my inputs', tmp_path / 'new results'
    source.mkdir()
    result = run('--input', source, '--output', output)
    assert result.returncode == 0, result.stderr
    args = json.loads(log.read_text().splitlines()[-1])
    assert f'{source}:/source:ro' in args
    assert f'{output}:/result:rw' in args
    assert args[-4:] == ['--input', '/source', '--output', '/result']


def test_existing_results_are_refused_without_running_docker(launcher):
    case, log, run = launcher
    result = run('--output', case / 'output')
    assert result.returncode == 2
    assert 'must be empty' in result.stderr
    assert not log.exists()
    assert (case / 'output/published.xlsx').read_bytes() == b'published result'


@pytest.mark.parametrize('relative', ['input/files', 'input/files/nested', 'input'])
def test_overlapping_folders_refused(launcher, relative):
    case, log, run = launcher
    result = run('--output', case / relative)
    assert result.returncode == 2
    assert 'separate, non-nested' in result.stderr
    assert not log.exists()


def test_symlink_into_inputs_refused(launcher):
    case, log, run = launcher
    alias = case / 'alias'
    alias.symlink_to(case / 'input/files', target_is_directory=True)
    result = run('--output', alias)
    assert result.returncode == 2
    assert not log.exists()


def test_refresh_trace_records_actual_invocations():
    trace = json.loads(Path('evidence/refresh-trace.json').read_text())
    events = trace['events']
    assert [e['rows'] for e in events] == [0, 5520, 5520, 5520, 5520, 5521]
    assert [e['phase'] for e in events] == ['invoke RefreshFolder', 'Complete'] * 3
    assert [e['elapsed_seconds'] for e in events] == sorted(e['elapsed_seconds'] for e in events)
    assert trace['original_sources_unchanged']
    assert trace['source_sha256'] == json.loads(Path('output/report.json').read_text())['source_sha256']
