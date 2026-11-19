"""Remove authoring-application properties without changing workbook data."""
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


def clean_properties(path):
    path=Path(path)
    with ZipFile(path) as z:
        items=[(entry,z.read(entry.filename)) for entry in z.infolist()]
    with ZipFile(path,'w') as z:
        for entry,data in items:
            if entry.filename.startswith('docProps/') and entry.filename.endswith('.xml'):
                root=ET.fromstring(data)
                for element in root.iter():
                    tag=element.tag.split('}')[-1]
                    if tag in ('Application','AppVersion','Company','Manager','Template'):
                        element.text=''
                    if tag in ('creator','lastModifiedBy'):
                        element.text='Genco Doganay'
                data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
            z.writestr(entry,data)
