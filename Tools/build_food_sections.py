"""Переносит разделы выписок в плейсхолдеры общего шаблона питания."""

import argparse
import copy
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

from food_run import ROOT, load_run
from food_docx_language import ukrainian_part

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
OPERATIONS = {'dry': 'сухпай', 'removal': 'снять с продовольствия', 'enrollment': 'зачислить на продовольствие'}
SLOTS = {'dry': 'GiveDryFood', 'removal': 'RemoveFromFoodSupport', 'enrollment': 'EnrollToFoodSupport'}


def text(paragraph):
    return ''.join(node.text or '' for node in paragraph.iter(W + 't'))


def build_sections(folder, output=None):
    run = load_run(folder)
    template = ROOT / 'Documents/Templates/Extracts [food]/Template_{mm}_{OrderNumber}.docx'
    target = Path(output) if output else Path(run['results']) / 'Extracts' / f'Template_{folder[5:7]}_{run["order"]["number"]}.docx'
    blocks = {}
    for group in run['groups']:
        own = group['military'] == 'А7383'
        token = '{' + SLOTS[group['kind']] + ('_Ours' if own else '_Others') + '}'
        belonging = '[НАШИ]' if own else '[ЧУЖИЕ]'
        paths = [p for p in (Path(run['results']) / 'Extracts').glob('*.docx') if OPERATIONS[group['kind']] in p.name and belonging in p.name and f'- {group["reporter"]})' in p.name]
        if len(paths) != 1:
            raise ValueError(f'Неоднозначная выписка: {group["reporter"]}, {group["kind"]}')
        with zipfile.ZipFile(paths[0]) as archive:
            tree = ET.fromstring(archive.read('word/document.xml'))
        paragraphs = tree.find(W + 'body').findall(W + 'p')
        first = next(i for i, p in enumerate(paragraphs) if text(p).startswith('Нижчепойменованих'))
        last = next(i for i, p in enumerate(paragraphs) if text(p).startswith('Підстава:'))
        chunk = [copy.deepcopy(p) for p in paragraphs[first:last + 1]]
        # Переносим заголовок раздела; герб, общую шапку приказа и подписи не дублируем.
        node = next(chunk[0].iter(W + 't'))
        node.text = 'NN.\u00a0' + (node.text or '')
        node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        for paragraph in chunk:
            for node in paragraph.iter(W + 't'):
                node.text = (node.text or '').translate(str.maketrans('«»“”„', '"""""'))
        blocks.setdefault(token, {}).setdefault(text(chunk[0]), []).append(chunk)
    with zipfile.ZipFile(template) as original:
        tree = ET.fromstring(original.read('word/document.xml'))
        body = tree.find(W + 'body')
        found = set()
        for paragraph in list(body.findall(W + 'p')):
            tokens = re.findall(r'\{(?:EnrollToFoodSupport|RemoveFromFoodSupport|GiveDryFood)_(?:Ours|Others)\}', text(paragraph))
            if not tokens:
                continue
            token, = tokens
            found.add(token)
            for header_index, chunks in enumerate(blocks.get(token, {}).values()):
                if header_index:
                    blank = copy.deepcopy(paragraph)
                    for child in list(blank):
                        if child.tag != W + 'pPr':
                            blank.remove(child)
                    paragraph.addprevious(blank)
                for index, chunk in enumerate(chunks):
                    # Общая шапка одна; принадлежность, даты и основания остаются раздельными.
                    for inserted in (chunk if index == 0 else chunk[1:]):
                        paragraph.addprevious(inserted)
            if token in blocks:
                body.remove(paragraph)
            else:
                # Пустая категория не должна оставлять в результате NN или плейсхолдер.
                for child in list(paragraph):
                    if child.tag != W + 'pPr':
                        paragraph.remove(child)
        assert set(blocks) <= found
        xml = ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        assert not re.search(r'\{[^{}]+\}', text(body))
        target.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as final:
            for item in original.infolist():
                final.writestr(item, ukrainian_part(item.filename, xml if item.filename == 'word/document.xml' else original.read(item.filename)))
    with zipfile.ZipFile(template) as original, zipfile.ZipFile(target) as final:
        assert original.namelist() == final.namelist()
        for item in original.namelist():
            if item != 'word/document.xml':
                assert ukrainian_part(item, original.read(item)) == final.read(item)
    print(target)
    return target


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--date', required=True)
    parser.add_argument('--output')
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')
    build_sections(args.date, args.output)