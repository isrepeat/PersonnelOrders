import collections
import hashlib
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).parent
SOURCE = Path(r'C:\Users\isrepeat\OneDrive\Рабочий стол\РУХ_last.xlsx')
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def normalize(value):
    return ' '.join(value.replace('\u00a0', ' ').replace('\u202f', ' ').split()).casefold()


digest = hashlib.sha256(SOURCE.read_bytes()).hexdigest()
with zipfile.ZipFile(SOURCE) as book:
    shared = ET.fromstring(book.read('xl/sharedStrings.xml'))
    strings = [''.join(t.text or '' for t in si.findall('.//s:t', NS)) for si in shared]
    workbook = ET.fromstring(book.read('xl/workbook.xml'))
    relationships = ET.fromstring(book.read('xl/_rels/workbook.xml.rels'))
    targets = {item.attrib['Id']: item.attrib['Target'] for item in relationships}
    result = {'path': str(SOURCE), 'sha256': digest}
    for kind, sheet_name, columns in [('list', 'Список', ['B', 'C']), ('arrivals', 'Прибулі', ['C', 'D', 'F'])]:
        sheet = next(item for item in workbook.find('s:sheets', NS) if item.attrib['name'] == sheet_name)
        target = targets[sheet.attrib['{' + REL + '}id']]
        target = target.lstrip('/') if target.startswith('/') else 'xl/' + target
        xml = ET.fromstring(book.read(target))
        rows = []
        for row in xml.findall('s:sheetData/s:row', NS):
            number = int(row.attrib['r'])
            if number < 6 or number > 20000:
                continue
            cells = {}
            for cell in row:
                address = cell.attrib.get('r', '')
                column = ''.join(c for c in address if c.isalpha())
                if column not in columns:
                    continue
                value = cell.find('s:v', NS)
                text = value.text if value is not None and value.text is not None else ''
                if cell.attrib.get('t') == 's':
                    text = strings[int(text)]
                elif cell.attrib.get('t') == 'inlineStr':
                    text = ''.join(t.text or '' for t in cell.findall('.//s:t', NS))
                elif cell.attrib.get('t') == 'e':
                    raise ValueError(f'Source cell error at {sheet_name}!{address}')
                cells[column] = text
            values = [cells.get(column, '') for column in columns]
            if any(values):
                rows.append({'row': number, 'name': values[0], 'id': values[1], 'origin': values[2] if kind == 'arrivals' else ''})
        result[kind] = rows
registry = Path(r'C:\WORK\Windows\Projects\ExcelPrototype\PROJECTS_EXCEL\workbook\PeopleRegistry\PeopleRegistry.xlsm')
with zipfile.ZipFile(registry) as book:
    shared = ET.fromstring(book.read('xl/sharedStrings.xml'))
    strings = [''.join(t.text or '' for t in si.findall('.//s:t', NS)) for si in shared]
    workbook = ET.fromstring(book.read('xl/workbook.xml'))
    relationships = ET.fromstring(book.read('xl/_rels/workbook.xml.rels'))
    targets = {item.attrib['Id']: item.attrib['Target'] for item in relationships}
    sheet = next(item for item in workbook.find('s:sheets', NS) if item.attrib['name'] == 'Історія ID')
    target = targets[sheet.attrib['{' + REL + '}id']]
    target = target.lstrip('/') if target.startswith('/') else 'xl/' + target
    xml = ET.fromstring(book.read(target))
    identifiers = []
    for cell in xml.findall('s:sheetData/s:row/s:c', NS):
        address = cell.attrib.get('r', '')
        if ''.join(c for c in address if c.isalpha()) not in ['A', 'B']:
            continue
        if int(''.join(c for c in address if c.isdigit())) < 4:
            continue
        value = cell.find('s:v', NS)
        text = value.text if value is not None and value.text is not None else ''
        if cell.attrib.get('t') == 's':
            text = strings[int(text)]
        if text:
            identifiers.append(text)
    result['reservedHistoryIDs'] = sorted(set(identifiers))
    history_records = []
    for row in xml.findall('s:sheetData/s:row', NS):
        if int(row.attrib['r']) < 4:
            continue
        cells = {}
        for cell in row:
            column = ''.join(c for c in cell.attrib.get('r', '') if c.isalpha())
            value = cell.find('s:v', NS)
            text = value.text if value is not None and value.text is not None else ''
            if cell.attrib.get('t') == 's':
                text = strings[int(text)]
            cells[column] = text
        if cells.get('A'):
            history_records.append({'id': cells['A'], 'currentID': cells.get('B', ''), 'name': cells.get('C', ''), 'origin': cells.get('D', '')})
    result['historyRecords'] = history_records
ROOT.joinpath('source.json').write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
names = collections.defaultdict(set)
for row in result['list']:
    names[normalize(row['name'])].add(row['id'])
groups = collections.Counter((normalize(row['name']), normalize(row['origin'])) for row in result['arrivals'])
print(json.dumps({'listRows': len(result['list']), 'arrivalRows': len(result['arrivals']), 'uniqueArrivalNameOrigin': len(groups), 'repeatedArrivalGroups': sum(n > 1 for n in groups.values()), 'repeatedArrivalRows': sum(n for n in groups.values() if n > 1), 'namesWithMultipleIDsInList': sum(len(ids) > 1 for ids in names.values()), 'arrivalsWithOneListID': sum(len(names.get(normalize(row['name']), set())) == 1 for row in result['arrivals']), 'arrivalsWithMultipleListIDs': sum(len(names.get(normalize(row['name']), set())) > 1 for row in result['arrivals']), 'arrivalsWithoutListID': sum(not names.get(normalize(row['name'])) for row in result['arrivals'])}))