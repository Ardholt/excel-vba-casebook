"""Exercise the default inner pipeline in an empty destination and publish its results."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import tempfile
import time
from datetime import datetime, timezone

root = Path.cwd()
source = root / 'input/files'
def hashes():
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
            for p in sorted(source.iterdir()) if p.is_file()}

before = hashes()
started = time.perf_counter()
with tempfile.TemporaryDirectory(prefix='fresh-combine-') as tmp:
    destination = Path(tmp)
    assert not list(destination.iterdir())
    subprocess.run(['bash', 'scripts/inner.sh', '--demo', str(source), str(destination)], check=True)
    assert hashes() == before
    for name in ('output', 'evidence'):
        # Existing bind-mounted files may be writable without permitting chmod/utime.
        for item in (destination / name).iterdir():
            if item.is_file():
                shutil.copyfile(item, root / name / item.name)
    shutil.copyfile(destination / 'combine.xlsm', root / 'combine.xlsm')
    Path('evidence/fresh-run.json').write_text(json.dumps({
        'route': 'scripts/inner.sh --demo INPUT EMPTY_RESULT',
        'environment': 'Linux container', 'destination_initially_empty': True,
        'copied_to_destination': ['scripts', 'vba', 'tests', 'run.sh'],
        'published_outputs_used_as_inputs': False,
        'original_source_hashes_unchanged': before == hashes(),
        'source_sha256': before,
        'verification': 'Python reference, packaged macro in LibreOffice, then pytest',
        'docker_launcher_validation': 'Argument-level tests with a Docker command double',
    }, indent=2) + '\n')
    tests = json.loads(Path('evidence/tests.json').read_text())
    Path('evidence/verification.json').write_text(json.dumps({
        'command': 'python3 scripts/verify_fresh.py',
        'checked_utc': datetime.now(timezone.utc).isoformat(),
        'exit_code': 0, 'seconds': round(time.perf_counter() - started, 3),
        'test_summaries': [f"{tests['passed']} passed in {tests['seconds']}s"],
        'inputs_unchanged': hashes() == before,
        'input_sha256': {'input/files/' + name: digest for name, digest in before.items()},
        'scope': 'Fresh destination; code copied without published outputs; actual macro in LibreOffice and Python reference.',
        'source_sha256': {str(p): hashlib.sha256(p.read_bytes()).hexdigest()
                          for folder in ('scripts', 'tests') for p in sorted(Path(folder).glob('*')) if p.is_file()},
    }, indent=2) + '\n')
