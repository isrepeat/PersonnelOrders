"""Сверяет таблицу и выписки Food с выбранной папкой Sources и шаблонами."""

import argparse
import re
import sys
import zipfile
from datetime import date, timedelta
from pathlib import Path

from docx import Document
from lxml import etree as ET
from openpyxl import load_workbook
from pypdf import PdfReader

from food_person_genitive import person_accusative
from food_run import ROOT, load_run

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def verify(folder, workbook):
    run = load_run(folder)
    book = load_workbook(workbook)
    cached = load_workbook(workbook, data_only=True)
    reference = load_workbook(run['reference'])
    assert book.sheetnames == ['Продовольче', 'Сухпрод']
    for name, kind, table_name, last_col in [('Продовольче', 'removal', 'FoodReports', 'O'), ('Сухпрод', 'dry', 'DryRations', 'N')]:
        sheet = book[name]
        expected = [(group, person) for group in run['groups'] if group['kind'] == kind for person in group['people']]
        assert sheet.tables[table_name].ref == f'B2:{last_col}{max(3, len(expected) + 2)}'
        assert sheet.freeze_panes is None and sheet.sheet_view.showGridLines is False
        assert sheet.column_dimensions['A'].style
        for row, (group, person) in enumerate(expected, 3):
            assert [sheet.cell(row, c).value for c in (2, 3, 4)] == [person['rank'], person['name'], group['military']]
            if kind == 'dry':
                assert sheet.cell(row, 5).value.date().isoformat() == group['start']
                assert sheet.cell(row, 6).value == group['duration']
                assert 'Початок' in sheet.cell(row, 7).value and 'Тривалість' in sheet.cell(row, 7).value
                assert cached[name].cell(row, 7).value.date() == date.fromisoformat(group['start']) + timedelta(days=group['duration'])
                basis_col, order_col, info_col, source_col, ipn_col = 8, 9, 10, 11, 14
            else:
                assert all(sheet.cell(row, c).value is None for c in (5, 6, 7))
                assert sheet.cell(row, 8).value.date().isoformat() == group['start']
                basis_col, order_col, info_col, source_col, ipn_col = 10, 9, 11, 12, 15
            assert sheet.cell(row, basis_col).value == group['basis']
            assert sheet.cell(row, order_col).value == run['order']['number']
            assert sheet.cell(row, info_col).value is None
            assert sheet.cell(row, source_col).value == '; '.join(group['source_pages'])
            assert sheet.cell(row, ipn_col).value in (person.get('ipn'), '')
            assert sheet.cell(row, 3).fill.fgColor.rgb[-6:] == ('262626' if row % 2 else '383838')
        for row in sheet:
            assert all(not (isinstance(cell.value, str) and cell.value.startswith(('#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A', '#NUM!'))) for cell in row)
    for col in 'ABCDEFGHIJ':
        assert book['Сухпрод'].column_dimensions[col].width == reference['Сухпрод'].column_dimensions[col].width
    rules = [rule for group in book['Сухпрод'].conditional_formatting._cf_rules.values() for rule in group]
    assert any(rule.type == 'duplicateValues' for rule in rules)
    assert book['Сухпрод'].tables['DryRations'].tableColumns[5].calculatedColumnFormula is not None
    for group in run['groups']:
        operation = 'сухпай' if group['kind'] == 'dry' else 'снять с продовольствия'
        belonging = '[НАШИ]' if group['military'] == 'А7383' else '[ЧУЖИЕ]'
        template = next(path for path in (ROOT / 'Documents/Templates/Extracts [food]').glob('*.docx') if operation in path.name and belonging in path.name)
        outputs = [path for path in (Path(run['results']) / 'Extracts').glob('*.docx') if f'- {group["reporter"]})' in path.name]
        corrected = [path for path in (Path(run['results']) / 'Extracts/Исправленные основания').glob('*.docx') if f'- {group["reporter"]})' in path.name]
        if corrected:
            outputs = corrected
        assert len(outputs) == 1
        output = outputs[0]
        assert output.name.startswith(f'{run["order"]["date"]} {run["order"]["number"]} ')
        with zipfile.ZipFile(template) as original, zipfile.ZipFile(output) as final:
            assert original.namelist() == final.namelist()
            for part in original.namelist():
                if part != 'word/document.xml':
                    assert original.read(part) == final.read(part), f'Изменена часть шаблона: {part}'
            source_tree, tree = ET.fromstring(original.read('word/document.xml')), ET.fromstring(final.read('word/document.xml'))
            assert ET.tostring(source_tree.find('.//w:sectPr', NS)) == ET.tostring(tree.find('.//w:sectPr', NS))
            text = '\n'.join(''.join(p.xpath('.//w:t/text()', namespaces=NS)) for p in tree.findall('.//w:p', NS))
            assert not re.search(r'\{[^{}]+\}', text)
            for number, person in enumerate(group['people'], 1):
                rank, name = person_accusative(person)
                punctuation = '.' if number == len(group['people']) else ';'
                assert f'{number}. {rank} {name}{punctuation}' in text
            for person in group.get('excluded_people', []):
                assert person['name'].split()[0] not in text
            assert re.sub(r'^лист\b', 'рапорт', group['raport']) in text and group['incoming_number'] in text
            assert run['order']['name'] in text and run['order']['rank'] in text
            if group.get('meal'):
                assert group['meal'] not in text
            # Исходные свойства абзацев и фрагментов остаются на месте; добавлен только список.
            source_doc, doc = Document(template), Document(output)
            source_props = [ET.tostring(p._p.pPr) if p._p.pPr is not None else b'' for p in source_doc.paragraphs]
            final_paragraphs = [p for p in doc.paragraphs if not re.match(r'^\d+\. ', p.text)]
            basis_index = next(i for i, p in enumerate(final_paragraphs) if p.text.startswith('Підстава:'))
            del final_paragraphs[basis_index - 1]
            assert len(final_paragraphs) == len(source_props)
            assert [ET.tostring(p._p.pPr) if p._p.pPr is not None else b'' for p in final_paragraphs] == source_props
        print(f'Выписка проверена: {output.name}')
    dry = [(group, person) for group in run['groups'] if group['kind'] == 'dry' for person in group['people']]
    missing = [person['name'] for _, person in dry if 'Не знайдено' in person['match']]
    changed_ranks = [person['name'] for _, person in dry if person['rank'] != person['source_rank']]
    print(f'{folder}: приказ №{run["order"]["number"]}; сухпай {len(dry)}; снятие {sum(len(g["people"]) for g in run["groups"] if g["kind"] == "removal")}; комплектов {sum(g["duration"] for g, _ in dry)}')
    print(f'Без совпадения ШПО: {len(missing)}; расхождения званий: {len(changed_ranks)}')
    for pdf in Path(run['work']).glob('final-*.pdf'):
        print(f'Word-рендер: {pdf.name}, страниц {len(PdfReader(pdf).pages)}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--date', required=True)
    parser.add_argument('--workbook', required=True)
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')
    verify(args.date, args.workbook)