"""Remove personal and application-identifying document properties."""
from pathlib import Path
from io import BytesIO
import zipfile
import xml.etree.ElementTree as ET


def scrub(path):
    path = Path(path)
    out = BytesIO()
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as dst:
        for entry in src.infolist():
            data = src.read(entry.filename)
            if entry.filename in ('docProps/core.xml', 'docProps/app.xml'):
                tree = ET.fromstring(data)
                for node in tree:
                    tag = node.tag.split('}')[-1]
                    if tag in ('creator', 'lastModifiedBy'):
                        node.text = 'Genco Doganay'
                    elif tag in ('Application', 'AppVersion', 'Company', 'Manager', 'HyperlinkBase', 'Template'):
                        node.text = ''
                data = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
            dst.writestr(entry, data)
    path.write_bytes(out.getvalue())

if __name__ == '__main__':
    import sys
    for name in sys.argv[1:]:
        scrub(name)
