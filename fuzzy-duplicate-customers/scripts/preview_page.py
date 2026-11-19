"""Capture the page offline, replacing only the network-only status badge."""
import json
import subprocess
from pathlib import Path
readme=Path('README.md').read_text()
badge='[![Tests](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg)](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)'
assert badge in readme
preview=Path('.page-preview.md')
try:
    preview.write_text(readme.replace(badge,'[Tests — online status](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)'))
    subprocess.run(['pagerender',str(preview),'evidence/page'],check=True)
finally:
    preview.unlink(missing_ok=True)
Path('evidence/page/capture-notes.json').write_text(json.dumps({'source':'README.md','offline_adjustment':'The remote CI status image is replaced by its text link for this offline capture. The published README retains the requested badge.','limitations':'This is a local page rendering, not a deployed GitHub or Pages session.'},indent=2)+'\n')
