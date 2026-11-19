from pathlib import Path
import zipfile
from pyopenvba import ExcelFile
Path('.build').mkdir(exist_ok=True)
parts=[]
for name in ['SplitRows','MergeTables','RenameSheets','MakeLetters']:
    parts.append('\n'.join(l for l in Path('vba',name+'.bas').read_text().splitlines() if not l.startswith('Attribute ') and l!='Option Explicit'))
source='Option Explicit\n'+'\n'.join(parts)+'\n'
with ExcelFile.create_new('.build/code.xlsm') as b:
    b.set_module('Module1',source); b.save()
with zipfile.ZipFile('.build/code.xlsm') as z: Path('.build/vbaProject.bin').write_bytes(z.read('xl/vbaProject.bin'))
Path('.build/embedded.bas').write_text(source)
