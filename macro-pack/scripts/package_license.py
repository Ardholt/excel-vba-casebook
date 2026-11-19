from pathlib import Path
import importlib.metadata as metadata
package=metadata.distribution('pyopenvba')
licenses=[p for p in package.files if 'license' in str(p).lower() and str(p).lower().endswith(('license','.txt','.md'))]
assert licenses, 'Packaging licence not found'
Path('evidence/packaging-LICENSE.txt').write_text('\n\n'.join(package.locate_file(p).read_text() for p in licenses))
print('Packaging version',package.version)
