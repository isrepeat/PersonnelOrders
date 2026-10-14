"""Добавляет выдачу аттестата в существующие выписки и общие разделы без пересборки текста."""

import argparse
import copy
import json
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree as ET

from food_docx_language import ukrainian_part

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
SENTENCE = 'Видати продовольчий атестат.'
parser = argparse.ArgumentParser()
parser.add_argument('--date', required=True)
args = parser.parse_args()
sys.stdout.reconfigure(encoding='utf-8')
task = Path(__file__).resolve().parents[1] / 'Tasks/01.Food'
groups = json.loads((task / 'Work' / args.date / 'reports.json').read_text(encoding='utf-8'))['groups']
requested = [group for group in groups if group.get('food_card')]
backup = task / 'Work' / args.date / 'Versions' / datetime.now().strftime('Before_food_card_%Y%m%d%H%M%S')
for path in (task / 'Results' / args.date / 'Extracts').glob('*.docx'):
    with zipfile.ZipFile(path) as archive:
        parts = [(item, archive.read(item.filename)) for item in archive.infolist()]
    tree = ET.fromstring(dict((item.filename, data) for item, data in parts)['word/document.xml'])
    paragraphs = tree.find(W + 'body').findall(W + 'p')
    changed = False
    for group in requested:
        for index, paragraph in enumerate(paragraphs):
            text = ''.join(paragraph.itertext())
            if not text.startswith('Підстава:') or group['incoming_number'] not in text:
                continue
            previous_basis = max((i for i in range(index) if ''.join(paragraphs[i].itertext()).startswith('Підстава:')), default=-1)
            if any(SENTENCE in ''.join(p.itertext()) for p in paragraphs[previous_basis + 1:index]):
                continue
            last = next(p for p in reversed(paragraphs[:index]) if ''.join(p.itertext()).strip())
            added = copy.deepcopy(last)
            for child in list(added):
                if child.tag != W + 'pPr':
                    added.remove(child)
            run = ET.SubElement(added, W + 'r')
            original_run = last.find(W + 'r')
            if original_run is not None and original_run.find(W + 'rPr') is not None:
                run.append(copy.deepcopy(original_run.find(W + 'rPr')))
            ET.SubElement(run, W + 't').text = SENTENCE
            last.addnext(added)
            changed = True
    if not changed:
        continue
    backup.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, backup / path.name)
    xml = ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    staged = backup / ('updated-' + path.name)
    with zipfile.ZipFile(staged, 'w') as archive:
        for item, data in parts:
            archive.writestr(item, ukrainian_part(item.filename, xml if item.filename == 'word/document.xml' else data))
    shutil.copy2(staged, path)
    print(path)