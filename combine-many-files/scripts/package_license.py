"""Retain the licence covering the embedded workbook template."""
from importlib.metadata import distribution
from pathlib import Path
package = distribution('pyopenvba')
licenses = [p for p in package.files if p.name.lower() in ('license','license.txt','license.md')]
if len(licenses) != 1:
    raise RuntimeError('Expected one workbook-template licence')
Path('evidence/PYOPENVBA-LICENSE.txt').write_text(package.locate_file(licenses[0]).read_text())
print('Retained workbook-template licence')
