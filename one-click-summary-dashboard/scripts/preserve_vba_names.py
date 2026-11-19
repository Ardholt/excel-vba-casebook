"""Keep VBA document module names when Calc exports an existing XLSM project."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
def preserve(path):
    path=Path(path)
    with ZipFile(path) as z: payload={n:z.read(n) for n in z.namelist()}
    root=ET.fromstring(payload['xl/workbook.xml'])
    props=root.find('{'+NS+'}workbookPr')
    if props is None:props=ET.SubElement(root,'{'+NS+'}workbookPr')
    props.set('codeName','ThisWorkbook')
    # Calc can retain the original print area as well as its updated one.
    names=root.find('{'+NS+'}definedNames')
    if names is not None:
        seen=set()
        for node in reversed(list(names)):
            key=(node.get('name'),node.get('localSheetId'))
            if key in seen:names.remove(node)
            seen.add(key)
    payload['xl/workbook.xml']=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    for i in range(1,5):
        name=f'xl/worksheets/sheet{i}.xml'
        sheet=ET.fromstring(payload[name])
        props=sheet.find('{'+NS+'}sheetPr')
        if props is None:
            props=ET.Element('{'+NS+'}sheetPr');sheet.insert(0,props)
        props.set('codeName',f'Sheet{i}')
        payload[name]=ET.tostring(sheet,encoding='utf-8',xml_declaration=True)
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for name,data in payload.items():z.writestr(name,data)

if __name__=='__main__':
    import sys
    preserve(sys.argv[1])
