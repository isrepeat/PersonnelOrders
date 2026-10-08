import sys
import re
import zipfile
import difflib
import xml.etree.ElementTree as ET

# Частичное окрашивание текста добавляется в OOXML после экспорта: API не поддерживает текстовые фрагменты.
path = sys.argv[1]
ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', ns)
q = lambda name: '{' + ns + '}' + name
with zipfile.ZipFile(path) as archive:
    files = {name: archive.read(name) for name in archive.namelist()}
sheet = ET.fromstring(files['xl/worksheets/sheet1.xml'])
shared = ET.fromstring(files['xl/sharedStrings.xml']) if 'xl/sharedStrings.xml' in files else None
cells = {cell.attrib['r']: cell for cell in sheet.iter(q('c'))}
def content(cell):
    if cell.attrib.get('t') == 'str':
        return cell.find(q('v')).text or ''
    if cell.attrib.get('t') == 's':
        node = shared[int(cell.find(q('v')).text)]
    else:
        node = cell.find(q('is'))
    return ''.join(node.itertext()) if node is not None else ''
count = 0
for address, cell in cells.items():
    if not address.startswith('M'):
        continue
    row = int(address[1:])
    if row < 11 or (row - 11) % 4 < 2:
        continue
    old = content(cells['M' + str(row - 2)])
    new = content(cell)
    if old == new and cell.attrib.get('t') != 'inlineStr':
        continue
    for child in list(cell):
        cell.remove(child)
    cell.set('t', 'inlineStr')
    inline = ET.SubElement(cell, q('is'))
    for kind, a, b, c, d in difflib.SequenceMatcher(None, old, new, autojunk=False).get_opcodes():
        if c == d:
            continue
        run = ET.SubElement(inline, q('r'))
        prop = ET.SubElement(run, q('rPr'))
        ET.SubElement(prop, q('rFont'), {'val': 'Arial'})
        ET.SubElement(prop, q('sz'), {'val': '10'})
        if kind != 'equal':
            ET.SubElement(prop, q('color'), {'rgb': 'FFFF4040'})
            ET.SubElement(prop, q('b'))
        else:
            ET.SubElement(prop, q('color'), {'rgb': 'FFF0ECE6'})
        text = ET.SubElement(run, q('t'), {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'})
        text.text = new[c:d]
    count += 1
files['xl/worksheets/sheet1.xml'] = ET.tostring(sheet, encoding='utf-8', xml_declaration=True)
with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
    for name, data in files.items():
        archive.writestr(name, data)
print('Red text cells:', count)