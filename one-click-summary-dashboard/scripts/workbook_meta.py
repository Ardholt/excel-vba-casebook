"""Remove identifying document properties without changing workbook cells or caches."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET


def scrub(path):
    path = Path(path)
    payload = {}
    with ZipFile(path) as archive:
        for member in archive.namelist():
            data = archive.read(member)
            if member in ('docProps/core.xml', 'docProps/app.xml'):
                root = ET.fromstring(data)
                for element in root.iter():
                    if element.tag.rsplit('}', 1)[-1] in ('creator','lastModifiedBy','Application','AppVersion','Company','Manager'):
                        element.text = ''
                data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            payload[member] = data
    with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
        for member, data in payload.items():
            archive.writestr(member, data)
