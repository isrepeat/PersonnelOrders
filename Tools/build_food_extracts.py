import copy
import json
import sys
from pathlib import Path
from datetime import date, timedelta
from docx import Document
from lxml import etree as ET
from openpyxl import load_workbook
from extract_order import load_shpo
from food_person_genitive import person_genitive

root = Path(__file__).resolve().parents[1]
task = root / 'Tasks/01.Food'
groups = json.loads((task / 'Work/2026.10.09/reports.json').read_text(encoding='utf-8'))
templates = root / 'Documents/Templates/Extracts [food]'
out = Path(sys.argv[1]) if len(sys.argv) > 1 else task / 'Results/2026.10.09/Extracts'
out.mkdir(parents=True, exist_ok=True)
orders_book = load_workbook(root / 'Tools/Накази.xlsx', data_only=True)
orders_sheet = orders_book['Накази']
orders = {row[0].date(): row for row in orders_sheet.iter_rows(min_col=10, max_col=14, min_row=3, values_only=True) if hasattr(row[0], 'date')}
alf, ranks, positions = load_shpo()
months = ['січня', 'лютого', 'березня', 'квітня', 'травня', 'червня', 'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня']
def period_text(start, finish):
    if (start.year, start.month) == (finish.year, finish.month):
        return f'з {start:%d} по {finish:%d} {months[finish.month - 1]} {finish.year} року'
    first_year = f' {start.year} року' if start.year != finish.year else ''
    return f'з {start:%d} {months[start.month - 1]}{first_year} по {finish:%d} {months[finish.month - 1]} {finish.year} року'
assert period_text(date(2026, 9, 29), date(2026, 10, 1)) == 'з 29 вересня по 01 жовтня 2026 року'
authors = {'МІЗЯК': ('Андрій', 'капітан', 'МІЗЯКА'), 'МОРОЗОВИЧ': ('Юрій', 'лейтенант', 'МОРОЗОВИЧА'), 'КОРНЕНКО': ('Юрій', 'майор', 'КОРНЕНКА')}
rank_genitive = {'капітан': 'капітана', 'лейтенант': 'лейтенанта', 'майор': 'майора'}
reporter_positions = {
    'КОРНЕНКО': 'командира зведеного підрозділу мобільного зв’язку 1 ОПЗ військової частини А0565',
    'МОРОЗОВИЧ': 'командира роти радіоелектронної боротьби',
    'МІЗЯК': 'командира 1 батальйону безпілотних систем',
}
def replace_placeholders(paragraph, replacements):
    # Меняем только текст плейсхолдеров, сохраняя все фрагменты и свойства шаблона.
    for token, replacement in replacements.items():
        nodes = paragraph._p.xpath('.//w:t')
        text = ''.join(node.text or '' for node in nodes)
        while token in text:
            start = text.index(token)
            finish = start + len(token)
            offset = 0
            for node in nodes:
                content = node.text or ''
                end = offset + len(content)
                if offset < finish and end > start:
                    prefix = content[:max(0, start - offset)]
                    suffix = content[max(0, finish - offset):]
                    node.text = prefix + (replacement if offset <= start < end else '') + suffix
                    if node.text.startswith(' ') or node.text.endswith(' '):
                        node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                offset = end
            text = ''.join(node.text or '' for node in nodes)
for g in groups:
    template = next(p for p in templates.glob('*.docx') if ('[ЧУЖИЕ]' if g['military'] == 'А0565' else '[НАШИ]') in p.name and 'сухпай' in p.name)
    doc = Document(template)
    start = date.fromisoformat(g['start'])
    finish = start + timedelta(days=g['duration'] - 1)
    order_date = date.fromisoformat(g.get('report_date', '2026-10-08'))
    order = orders[order_date]
    if any(value is None for value in order[1:]):
        raise ValueError(f'Неполные реквизиты tbOrders: {start}')
    first_name, author_rank, surname_genitive = authors[g['reporter']]
    matches = [(ipn, name) for ipn, name in alf.items() if name.split()[:2] == [g['reporter'], first_name]]
    initials = first_name[0] + '.'
    if len(matches) == 1:
        initials = ''.join(word[0] + '.' for word in matches[0][1].split()[1:])
        author_rank = ranks.get(matches[0][0], author_rank).lower()
    author_text = f'{surname_genitive}\u00a0{initials}' if len(matches) == 1 else f"{ {'Юрій': 'Юрія', 'Андрій': 'Андрія'}.get(first_name, first_name)} {surname_genitive}"
    replacements = {
        '{dd}': f'{order_date:%d}', '{mm}': f'{order_date:%m}', '{yyyy}': f'{order_date:%Y}',
        '{OrderNumber}': str(order[1]), '{SignatoryRank}': order[4],
        '{SignatoryName}': order[3], '{Military}': g['military'],
        '{DurationText}': f"{g['duration']} доби", '{Period}': period_text(start, finish),
        '{Raport}': f"рапорт {reporter_positions[g['reporter']]} {rank_genitive[author_rank]} {author_text}",
        '{RaportNumber}': {'МІЗЯК': '1656/26790-в', 'МОРОЗОВИЧ': '1656/26786-в', 'КОРНЕНКО': '1656/26791-в'}.get(g['reporter'], '[ВХІДНИЙ НОМЕР]'),
        '{RaportDate}': '[ДАТА РЕЄСТРАЦІЇ]' if g['reporter'] == 'КОСЬКО' else '08.10.2026'
    }
    paragraphs = list(doc.paragraphs) + [p for t in doc.tables for row in t.rows for c in row.cells for p in c.paragraphs]
    for p in paragraphs:
        replace_placeholders(p, replacements)
    anchor = next(p for p in doc.paragraphs if p.text.startswith('Підстава:'))
    separator = copy.deepcopy(anchor._p.getprevious())
    prototype = next(p for p in doc.paragraphs if p.text.startswith('з ') and 'року' in p.text)
    for n, person in enumerate(g['people'], 1):
        element = copy.deepcopy(prototype._p)
        for child in list(element):
            if child.tag != '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr':
                element.remove(child)
        run = copy.deepcopy(next(r._r for r in prototype.runs if r.text))
        for child in list(run):
            if child.tag != '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr':
                run.remove(child)
        from docx.oxml import OxmlElement
        text = OxmlElement('w:t')
        punctuation = '.' if n == len(g['people']) else ';'
        person_rank, person_name = person_genitive(person)
        text.text = f"{n}. {person_rank} {person_name}{punctuation}"
        run.append(text)
        element.append(run)
        anchor._p.addprevious(element)
    # Исходный пустой абзац перед списком остаётся; такой же отделяет список от основания.
    anchor._p.addprevious(separator)
    filename = template.name
    filename_values = {
        '{yyyy}': f'{order_date:%Y}', '{mm}': f'{order_date:%m}', '{dd}': f'{order_date:%d}',
        '{OrderNumber}': str(order[1]), '{Raporter}': g['reporter'], '{Signatory}': order[2],
    }
    for token, value in filename_values.items():
        filename = filename.replace(token, value)
    # Для старых внешних шаблонов сохраняем автора до обновления их имени пользователем.
    if '{Raporter}' not in template.name:
        filename = filename.removesuffix('.docx') + f" {g['reporter']}.docx"
    doc.save(out / filename)

# Сборщик выписок не изменяет таблицы и их историю версий.
sys.exit(0)