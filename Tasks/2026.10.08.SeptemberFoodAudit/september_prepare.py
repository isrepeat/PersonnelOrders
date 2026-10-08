import collections
import datetime as dt
import gzip
import json
import re
from pathlib import Path

import openpyxl
from docx import Document

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent / 'Results'
analysis = json.load(gzip.open(OUT / 'analysis.json.gz', 'rt', encoding='utf-8'))
statistics = json.load(gzip.open(OUT / 'statistics.json.gz', 'rt', encoding='utf-8'))
book = openpyxl.load_workbook(OUT / 'Продовольче.xlsm', read_only=True, data_only=True)
sheet = book['Котел']
headers = list(next(sheet.iter_rows(min_row=2, max_row=2, max_col=47, values_only=True)))
positions = {name: i for i, name in enumerate(headers)}
documents = {}
assignments = []


def subdivision(text, unit):
    text = re.sub(r'\s+', ' ', text.lower())
    if unit != 'А7383':
        return 'СП.ЗІЧ'
    patterns = [
        ('шквал', 'СП. ШКВАЛ'),
        ('1 стрілецького батальйону', 'СП.1СБ'),
        ('стрілецького батальйону', 'СП.СБ'),
        ('1 механізованого батальйону', 'СП.1МБ'),
        ('2 механізованого батальйону', 'СП.2МБ'),
        ('2 танкового батальйону', 'СП.2ТБ'),
        ('танкового батальйону', 'СП.ТБ'),
        ('2 батальйону безпілотних систем', 'СП.2ББПС'),
        ('1 батальйону безпілотних систем', 'СП.ББПС'),
        ('самохідного артилерійського дивізіону', 'СП.САДН'),
        ('зенітного ракетного артилерійського дивізіону', 'СП.ЗРАДН'),
        ('ремонтно-відновлювального батальйону', 'СП.РВБ'),
        ('батальйону матеріального забезпечення', 'СП.БМЗ'),
        ('медичної роти', 'СП.МР'),
    ]
    for phrase, column in patterns:
        if phrase in text:
            return column
    return 'СП.ІП'


for group in analysis['groups']:
    start = dt.date.fromisoformat(group['start'])
    ambiguous = False
    working_start = start
    if working_start >= dt.date(2026, 10, 1) or working_start + dt.timedelta(days=group['days']) <= dt.date(2026, 9, 1):
        continue
    if group['file'] not in documents:
        documents[group['file']] = Document(Path(__file__).resolve().parent / 'Inputs/Orders' / group['file'])
    for person in group['people']:
        paragraph = documents[group['file']].paragraphs[person['paragraph'] - 1].text
        text = paragraph if group['order'] == 279 else group['heading'] + ' ' + group['basis']
        assignments.append({'name': person['name'], 'start': start.isoformat(), 'workingStart': working_start.isoformat(), 'days': group['days'], 'column': subdivision(text, group['unit']), 'order': group['order'], 'file': group['file'], 'paragraph': person['paragraph'], 'sourceText': paragraph, 'basis': group['basis'], 'ambiguous': ambiguous})

days = []
source_rows = {}
for rn, values in enumerate(sheet.iter_rows(min_row=3, max_col=47, values_only=True), 3):
    if isinstance(values[0], dt.datetime) and values[0].year == 2026 and values[0].month == 9:
        source_rows[values[0].date().isoformat()] = (rn, list(values))
for stats in statistics:
    day = dt.date.fromisoformat(stats['date'])
    rn, old = source_rows[stats['date']]
    old[0] = (day - dt.date(1899, 12, 30)).days
    new = old[:]
    for header, field in [('Продовольче', 'food'), ('Котлове', 'cooked'), ('Сухе', 'dry'), ('Знято', 'removed'), ('Зараховано', 'enrolled'), ('ЗІЧ', 'other')]:
        number = stats[field]
        assert isinstance(number, (int, float)) and number >= 0, (stats['date'], field, number)
        new[positions[header]] = int(number)
    assert stats['food'] == stats['cooked'] + stats['dry']
    new[positions['А7383+ТП']] = int(stats['food'] - stats['other'])
    allocation_known = any(old[i] is not None for i in range(8, 24))
    new[positions['ІП']] = new[positions['А7383+ТП']] - sum(old[i] or 0 for i in range(8, 24)) if allocation_known else None
    active = [a for a in assignments if dt.date.fromisoformat(a['workingStart']) <= day < dt.date.fromisoformat(a['workingStart']) + dt.timedelta(days=a['days'])]
    strict = [a for a in assignments if dt.date.fromisoformat(a['start']) <= day < dt.date.fromisoformat(a['start']) + dt.timedelta(days=a['days'])]
    counts = collections.Counter(a['column'] for a in active)
    new[positions['СП. ВСЬОГО']] = len(active)
    for col in range(26, 43):
        expected = counts[headers[col]]
        new[col] = expected if expected or old[col] is not None else None
    changes = [i + 1 for i in range(1, 43) if old[i] != new[i] and not (i >= 26 and not old[i] and not new[i])]
    notes = []
    if day.day <= 10:
        notes.append('Заполнены отсутствовавшие показатели РУХ; устранён двойной счёт сухпая. Разбивка продовольчего по подразделениям не восстановлена; ІП оставлено пустым.')
    if day.day >= 11:
        notes.append('Продовольче по актуальному РУХ на 1 меньше сохранённого значения; зависимые Котлове, А7383+ТП и ІП пересчитаны.')
    ambiguous = any(a['ambiguous'] for a in active)
    if ambiguous:
        notes.append(f'№257: в тексте 05–07 августа; 50 человек сохранены на сентябрь условно, как в обеих таблицах. Буквально по приказам на эту дату: {len(strict)}. Месяц требует подтверждения.')
    if day.day == 27:
        notes.append('№279: 11 человек относятся к СБ, а не 2ТБ; 2 человека — МР, а не ІП. Всего сухпая по приказам остаётся 75.')
    if int(stats['dry']) != len(active):
        notes.append(f'Расхождение источников: РУХ {int(stats["dry"])}, приказы {len(active)}; требуется отдельная проверка.')
    days.append({'date': stats['date'], 'sourceRow': rn, 'old': old, 'new': new, 'changes': changes, 'notes': ' '.join(notes), 'ambiguous': ambiguous, 'strictOrders': len(strict), 'workingOrders': len(active), 'stats': stats, 'orderNumbers': sorted({a['order'] for a in active})})
source_results = [row[0] for row in sheet.iter_rows(min_row=3, max_row=457, min_col=45, max_col=45, values_only=True)]
book.close()
assert len(days) == 30
payload = {'headers': headers, 'days': days, 'sourceResults': source_results, 'assignments': assignments, 'warnings': analysis['warnings'], 'summary': {'changedDays': sum(bool(d['changes']) for d in days), 'changedCells': sum(len(d['changes']) for d in days), 'ambiguousDays': sum(d['ambiguous'] for d in days)}}
with gzip.open(OUT / 'september.json.gz', 'wt', encoding='utf-8') as f:
    json.dump(payload, f, ensure_ascii=False)
print(json.dumps(payload['summary'], ensure_ascii=False))
print(json.dumps([{'date': d['date'], 'changes': [headers[i - 1] for i in d['changes']], 'dry': d['stats']['dry'], 'orders': d['workingOrders']} for d in days if d['date'] in ('2026-09-14', '2026-09-27', '2026-09-28')], ensure_ascii=False))