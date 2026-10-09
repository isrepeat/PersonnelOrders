import copy
import re
import sys
import zipfile
from pathlib import Path
from lxml import etree as ET

root = Path(__file__).resolve().parents[1]
folder = root / 'Tasks/01.Food/Results/2026.10.09'
versions = [(int(re.match(r'v(\d+) ', p.name)[1]), p) for p in folder.glob('v* Рапорти_продовольче_сухпрод_2026.10.09.xlsx')]
version, source = max(versions)
output = folder / f'v{version + 1} Рапорти_продовольче_сухпрод_2026.10.09.xlsx'
with zipfile.ZipFile(source) as archive:
    files = {name: archive.read(name) for name in archive.namelist()}
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
tag = lambda name: '{' + ns['s'] + '}' + name
styles = ET.fromstring(files['xl/styles.xml'])
fills = styles.find('s:fills', ns)
palette = {}
for color in ['262626', '383838', '000000']:
    palette[color] = len(fills)
    fill = ET.SubElement(fills, tag('fill'))
    pattern = ET.SubElement(fill, tag('patternFill'), patternType='solid')
    ET.SubElement(pattern, tag('fgColor'), rgb='FF' + color)
    ET.SubElement(pattern, tag('bgColor'), indexed='64')
fills.set('count', str(len(fills)))
fonts = styles.find('s:fonts', ns)
xfs = styles.find('s:cellXfs', ns)
cache = {}
def striped_style(original, fill_color, text_color):
    key = (original, fill_color, text_color)
    if key in cache:
        return cache[key]
    xf = copy.deepcopy(xfs[original])
    font = copy.deepcopy(fonts[int(xf.get('fontId', '0'))])
    old_color = font.find('s:color', ns)
    if old_color is not None:
        font.remove(old_color)
    ET.SubElement(font, tag('color'), rgb='FF' + text_color)
    xf.set('fontId', str(len(fonts)))
    fonts.append(font)
    xf.set('fillId', str(palette[fill_color]))
    xf.set('applyFill', '1')
    xf.set('applyFont', '1')
    cache[key] = len(xfs)
    xfs.append(xf)
    return cache[key]
for name in files:
    if not (name.startswith('xl/worksheets/sheet') and name.endswith('.xml')):
        continue
    sheet = ET.fromstring(files[name])
    for row in sheet.find('s:sheetData', ns):
        number = int(row.get('r'))
        for cell in row:
            col = re.match(r'[A-Z]+', cell.get('r'))[0]
            if col == 'A' or number < 2:
                continue
            fill = '000000' if number == 2 else ('262626' if number % 2 else '383838')
            text = 'FFFFFF'
            if number > 2 and col == 'D':
                text = 'FF9966'
            elif name.endswith('sheet2.xml') and number > 2 and col == 'G':
                text = '8DB4E2'
            cell.set('s', str(striped_style(int(cell.get('s', '0')), fill, text)))
    files[name] = ET.tostring(sheet, xml_declaration=True, encoding='UTF-8', standalone=True)
fonts.set('count', str(len(fonts)))
xfs.set('count', str(len(xfs)))
files['xl/styles.xml'] = ET.tostring(styles, xml_declaration=True, encoding='UTF-8', standalone=True)
with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
    for name, content in files.items():
        archive.writestr(name, content)
sys.stdout.reconfigure(encoding='utf-8')
print(output)