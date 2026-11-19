from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
from io import BytesIO
from PIL import Image
for p in list(Path('.').rglob('*.xlsx'))+list(Path('.').rglob('*.xlsm')):
    if '.build' in p.parts: continue
    output=BytesIO()
    with ZipFile(p) as source, ZipFile(output,'w',ZIP_DEFLATED) as target:
        for item in source.infolist():
            data=source.read(item.filename)
            if item.filename in ['docProps/core.xml','docProps/app.xml']:
                root=ET.fromstring(data)
                for node in root.iter():
                    if node.tag.split('}')[-1] in ['creator','lastModifiedBy','Application','AppVersion','Company','Manager']: node.text=''
                data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
            target.writestr(item,data)
    p.write_bytes(output.getvalue())
for p in Path('images').rglob('*.png'):
    im=Image.open(p); im.save(p)
