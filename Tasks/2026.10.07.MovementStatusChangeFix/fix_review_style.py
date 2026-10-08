import sys
import copy
import re
import zipfile
import xml.etree.ElementTree as E

ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
E.register_namespace('', ns)
q = lambda name: '{' + ns + '}' + name
with zipfile.ZipFile(sys.argv[1]) as archive:
    files = {name: archive.read(name) for name in archive.namelist()}
styles = E.fromstring(files['xl/styles.xml'])
formats = styles.find(q('numFmts'))
number_id = max([163] + [int(x.get('numFmtId')) for x in formats]) + 1
E.SubElement(formats, q('numFmt'), {'numFmtId': str(number_id), 'formatCode': 'dd.mm.yyyy'})
formats.set('count', str(len(formats)))
fills = styles.find(q('fills'))
fill = E.SubElement(fills, q('fill'))
pattern = E.SubElement(fill, q('patternFill'), {'patternType': 'solid'})
E.SubElement(pattern, q('fgColor'), {'rgb': 'FF292725'})
E.SubElement(pattern, q('bgColor'), {'indexed': '64'})
fill_id = str(len(fills) - 1)
fills.set('count', str(len(fills)))
fonts = styles.find(q('fonts'))
font = copy.deepcopy(fonts[0])
for color in font.findall(q('color')):
    font.remove(color)
E.SubElement(font, q('color'), {'rgb': 'FFF0ECE6'})
fonts.append(font)
fonts.set('count', str(len(fonts)))
xfs = styles.find(q('cellXfs'))
cache = {}
def new_style(old, top=False):
    key = (old, top)
    if key not in cache:
        xf = copy.deepcopy(xfs[old])
        if top:
            xf.set('fillId', fill_id)
            xf.set('fontId', str(len(fonts)-1))
            xf.set('applyFill', '1')
            xf.set('applyFont', '1')
        else:
            xf.set('numFmtId', str(number_id))
            xf.set('applyNumberFormat', '1')
        xfs.append(xf)
        cache[key] = str(len(xfs)-1)
    return cache[key]
sheet = E.fromstring(files['xl/worksheets/sheet2.xml'])
data = sheet.find(q('sheetData'))
for row in data:
    index = int(row.get('r'))
    if index <= 4:
        cells = {c.get('r'): c for c in row.findall(q('c'))}
        for column in 'ABCDEFGHIJKLM':
            address = column + str(index)
            cell = cells.get(address)
            if cell is None:
                cell = E.SubElement(row, q('c'), {'r': address})
            cell.set('s', new_style(int(cell.get('s', '0')), True))
        row[:] = sorted(row, key=lambda c: c.get('r'))
    else:
        for cell in row.findall(q('c')):
            if index >= 6 and cell.get('r')[0] in 'FGHI':
                cell.set('s', new_style(int(cell.get('s', '0'))))
xfs.set('count', str(len(xfs)))
def serialize(root, original):
    result = E.tostring(root, encoding='unicode')
    old_tag = re.search(r'<(?:\w+:)?(?:styleSheet|worksheet)\b[^>]*>', original.decode('utf-8'))[0]
    new_tag = result[:result.index('>')+1]
    missing = []
    for declaration in re.findall(r'xmlns(?::[\w]+)?="[^"]+"', old_tag):
        name = declaration.split('=')[0]
        if name + '=' not in new_tag:
            missing.append(declaration)
    result = result.replace(new_tag, new_tag[:-1] + (' ' + ' '.join(missing) if missing else '') + '>', 1)
    return result.encode('utf-8')
files['xl/styles.xml'] = serialize(styles, files['xl/styles.xml'])
files['xl/worksheets/sheet2.xml'] = serialize(sheet, files['xl/worksheets/sheet2.xml'])
with zipfile.ZipFile(sys.argv[2], 'w', zipfile.ZIP_DEFLATED) as archive:
    for name, content in files.items():
        archive.writestr(name, content)