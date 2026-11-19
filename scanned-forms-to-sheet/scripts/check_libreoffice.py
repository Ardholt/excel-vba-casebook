"""Compare a Calc-opened/exported workbook with the checked original."""
import json,sys,time,subprocess,tempfile
from pathlib import Path
import openpyxl
start=time.perf_counter()
checks={}
with tempfile.TemporaryDirectory() as temp:
    for name in ('input-inventory.xlsx','extracted.xlsx','reviewed.xlsx'):
        exported=Path(temp)/name
        subprocess.run(['xlrecalc',str(Path('output')/name),str(exported)],check=True,capture_output=True)
        source=openpyxl.load_workbook(Path('output')/name,data_only=True)
        converted=openpyxl.load_workbook(exported,data_only=True)
        same=source.sheetnames==converted.sheetnames
        for sheet in source.sheetnames:
            same=same and list(source[sheet].values)==list(converted[sheet].values)
        checks[name]=same
        assert same,name+' changed during Calc round trip'
runtime=json.loads(Path('output/runtime.json').read_text())
runtime['libreoffice_checks']={'method':'Opened and exported using LibreOffice Calc; every sheet value compared with the original using openpyxl.','workbooks':checks,'seconds':round(time.perf_counter()-start,3),'proof':'images/01-Proof.png rendered from output/proof.xlsx by LibreOffice Calc and Poppler.'}
Path('evidence/runtime.json').write_text(json.dumps(runtime,indent=2)+'\n')
print('PASS: LibreOffice round trip preserved all workbook values')
