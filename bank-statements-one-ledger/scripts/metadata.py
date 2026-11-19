"""Remove identifying author, company and application document properties."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET


def clean_workbook(path):
    path=Path(path); target=path.with_name('.'+path.name+'.clean')
    with ZipFile(path) as source, ZipFile(target,'w',ZIP_DEFLATED) as dest:
        for item in source.infolist():
            content=source.read(item.filename)
            if item.filename in ('docProps/core.xml','docProps/app.xml'):
                root=ET.fromstring(content)
                for el in root:
                    if el.tag.rsplit('}',1)[-1] in ('creator','lastModifiedBy','Application','AppVersion','Company','Manager','HyperlinkBase'):
                        el.text=''
                content=ET.tostring(root,encoding='utf-8',xml_declaration=True)
            dest.writestr(item,content)
    target.replace(path)
