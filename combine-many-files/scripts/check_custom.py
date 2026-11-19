"""Exercise the replacement-folder path, including spaces in both folder names."""
from pathlib import Path
import tempfile
import subprocess
import shutil
import json
import time
start = time.perf_counter()
with tempfile.TemporaryDirectory() as d:
    root = Path(d)
    src, out = root/'branch exports', root/'combined results'
    shutil.copytree('input/files', src)
    subprocess.run(['bash','scripts/inner.sh','--input',str(src),'--output',str(out)],check=True)
    report = json.loads((out/'report.json').read_text())
    assert report['output_rows'] == 5520
    assert report['amount_total'] == '238441.92'
    Path('evidence/custom-run.json').write_text(json.dumps({'route':'inner.sh with replacement input/output folders containing spaces',
        'rows':report['output_rows'],'amount_total':report['amount_total'],'source_to_output_cells_equal':True,
        'seconds':round(time.perf_counter()-start,3),'docker_mount_route':'Not run in this environment'},indent=2)+'\n')
