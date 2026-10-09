import json
import re
import zipfile
from pathlib import Path

from lxml import etree as ET
from food_person_genitive import person_genitive

root = Path(__file__).resolve().parents[1]
task = root / 'Tasks/01.Food'
groups = json.loads((task / 'Work/2026.10.09/reports.json').read_text(encoding='utf-8'))
folder = task / 'Results/2026.10.09/Extracts'
word = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
count = 0
for group in groups:
    paths = list(folder.rglob(f'2026* - {group["reporter"]}) *.docx'))
    for path in paths:
        with zipfile.ZipFile(path) as archive:
            files = {name: archive.read(name) for name in archive.namelist()}
        document = ET.fromstring(files['word/document.xml'])
        changed = 0
        for paragraph in document.findall('.//w:body/w:p', word):
            nodes = paragraph.findall('.//w:t', word)
            text = ''.join(node.text or '' for node in nodes)
            match = re.match(r'^(\d+)\. ', text)
            if not match:
                continue
            index = int(match[1])
            person = group['people'][index - 1]
            rank, name = person_genitive(person)
            punctuation = '.' if index == len(group['people']) else ';'
            replacement = f'{index}. {rank} {name}{punctuation}'
            nodes[0].text = replacement
            for node in nodes[1:]:
                node.text = ''
            changed += 1
        assert changed == len(group['people']), (path, changed)
        files['word/document.xml'] = ET.tostring(document, xml_declaration=True, encoding='UTF-8', standalone=True)
        try:
            target = path
            archive = zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED)
        except PermissionError:
            target = folder / 'Исправленные основания' / path.name
            target.parent.mkdir(exist_ok=True)
            archive = zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED)
        with archive:
            for name, content in files.items():
                archive.writestr(name, content)
        count += changed
        print(target)
print(f'Updated roster entries: {count}')