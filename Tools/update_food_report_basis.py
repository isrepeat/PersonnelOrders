import sys
import zipfile
from pathlib import Path
from lxml import etree as ET

root = Path(__file__).resolve().parents[1]
folder = root / 'Tasks/01.Food/Results/2026.10.09'
source = folder / 'v3 Рапорти_продовольче_сухпрод_2026.10.09.xlsx'
output = folder / 'v4 Рапорти_продовольче_сухпрод_2026.10.09.xlsx'
if output.exists():
    raise FileExistsError(output)
with zipfile.ZipFile(source) as archive:
    files = {name: archive.read(name) for name in archive.namelist()}
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
tag = lambda name: '{' + ns['s'] + '}' + name
sheet = ET.fromstring(files['xl/worksheets/sheet2.xml'])
bases = {
    'КОРНЕНКО': 'рапорт майора КОРНЕНКА Ю. (вх. № 1656/26791-в від 08.10.2026)',
    'МОРОЗОВИЧ': 'рапорт лейтенанта МОРОЗОВИЧА Ю.М. (вх. № 1656/26786-в від 08.10.2026)',
    'МІЗЯК': 'рапорт капітана МІЗЯКА А.В. (вх. № 1656/26790-в від 08.10.2026)',
}
for row in sheet.find('s:sheetData', ns):
    number = int(row.get('r'))
    if not 3 <= number <= 101:
        continue
    reporter = 'КОРНЕНКО' if number <= 78 else ('МОРОЗОВИЧ' if number <= 81 else 'МІЗЯК')
    for cell in row:
        if cell.get('r') == f'H{number}':
            for child in list(cell):
                cell.remove(child)
            cell.set('t', 'inlineStr')
            ET.SubElement(ET.SubElement(cell, tag('is')), tag('t')).text = bases[reporter]
        elif cell.get('r') == f'J{number}':
            for child in list(cell):
                cell.remove(child)
            cell.attrib.pop('t', None)
files['xl/worksheets/sheet2.xml'] = ET.tostring(sheet, xml_declaration=True, encoding='UTF-8', standalone=True)
with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
    for name, content in files.items():
        archive.writestr(name, content)
sys.stdout.reconfigure(encoding='utf-8')
print(output)