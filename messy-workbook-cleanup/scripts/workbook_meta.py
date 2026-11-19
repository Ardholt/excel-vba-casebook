"""Keep shipped workbook properties free of application and personal metadata."""
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED


def scrub_properties(path):
    with ZipFile(path) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    for name in ('docProps/core.xml', 'docProps/app.xml'):
        if name not in entries:
            continue
        root = ET.fromstring(entries[name])
        for node in root.iter():
            local = node.tag.rsplit('}', 1)[-1]
            if local in ('creator', 'lastModifiedBy'):
                node.text = 'Genco Doganay'
            if local in ('Application', 'AppVersion', 'Company', 'Manager', 'HyperlinkBase'):
                node.text = ''
        entries[name] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)
