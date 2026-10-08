import collections
import datetime as dt
import gzip
import json
from pathlib import Path

OUT = Path(__file__).resolve().parents[2] / 'outputs/dry-rations-audit-20261008'
p = json.load(gzip.open(OUT / 'analysis.json.gz', 'rt', encoding='utf-8'))
ALIASES = {'ЧМАЛ': 'ЧМАЛА', 'ІВАНЬОХ': 'ІВАНЬОХА', 'ЛЮБЧЕЦК': 'ЛЮБЧЕЦКО', 'ЛЮБКО': 'ЛЮБКА', 'ПОПЛИКА': 'ПОПЛИКО', 'КОСТИРОК': 'КОСТИРКО', 'БУЧКО': 'БУЧОК'}


def name_key(value):
    words = str(value).upper().replace('’', '').replace('`', '').replace("'", '').split()
    if words:
        words[0] = ALIASES.get(words[0], words[0])
    return ' '.join(words)


def unit_key(value):
    return str(value).upper().replace('A', 'А').strip()


events = []
for g in p['groups']:
    for person in g['people']:
        events.append({**person, **{k: g[k] for k in ('file', 'order', 'start', 'days', 'unit', 'basis')}, 'heading': g['heading'], 'sequence': len(events)})
        # Не создаём дубли из-за спорного месяца; имеющуюся дату оставляем для проверки.
        if g['order'] == 257 and g['start'] == '2026-08-05':
            events[-1]['sourceStart'] = g['start']
            events[-1]['start'] = '2026-09-05'
            events[-1]['note'] += ' Месяц требует подтверждения: приказ 04.09 указывает август, таблица — сентябрь. Дата таблицы сохранена.'
rows = [{'values': r['values'][:], 'sourceRow': r['row'], 'status': 'Вне периода проверки' if str(r['values'][3]) < '2026-08-01' else 'Не найдено в проверенных приказах', 'event': None, 'changes': ''} for r in p['existing']]
used = set()
unmatched = []
for event in events:
    candidates = [i for i, r in enumerate(rows) if i not in used and name_key(r['values'][1]) == name_key(event['name']) and r['values'][3] == event['start'] and unit_key(r['values'][2]) == unit_key(event['unit'])]
    exact = [i for i in candidates if rows[i]['values'][4] == event['days']]
    selected = exact[0] if exact else candidates[0] if len(candidates) == 1 else None
    if selected is None:
        unmatched.append(event)
        continue
    used.add(selected)
    r = rows[selected]
    old = r['values'][:]
    end = (dt.date.fromisoformat(event['start']) + dt.timedelta(days=event['days'])).isoformat()
    r['values'][:8] = [event['rank'], old[1], event['unit'], event['start'], event['days'], end, event['basis'], event['order']]
    r['event'] = event
    changes = [f'{label}: {old[j]} → {r["values"][j]}' for j, label in ((0, 'Звание'), (4, 'Длительность'), (5, 'Прекращение')) if old[j] != r['values'][j]]
    r['status'] = 'Исправлено по приказу' if changes else 'Подтверждено; заполнены основание и приказ'
    if event.get('sourceStart'):
        r['status'] = 'Требует проверки месяца; дата таблицы сохранена'
    r['changes'] = '; '.join(changes)

# После точных совпадений проверяем взаимно однозначные соседние даты.
# Повторные назначения и периоды дальше недели без прямой связи не объединяем.
def date_candidates(event):
    result = []
    for i, row in enumerate(rows):
        values = row['values']
        if i in used or name_key(values[1]) != name_key(event['name']) or unit_key(values[2]) != unit_key(event['unit']) or values[4] != event['days']:
            continue
        try:
            distance = abs((dt.date.fromisoformat(values[3]) - dt.date.fromisoformat(event['start'])).days)
        except (TypeError, ValueError):
            continue
        if 0 < distance <= 7:
            result.append(i)
    return result


candidate_map = [date_candidates(event) for event in unmatched]
references = collections.Counter(i for candidates in candidate_map for i in candidates)
still_unmatched = []
for event, candidates in zip(unmatched, candidate_map):
    if len(candidates) != 1 or references[candidates[0]] != 1:
        still_unmatched.append(event)
        continue
    i = candidates[0]
    used.add(i)
    row = rows[i]
    row['event'] = event
    row['values'][6:8] = [event['basis'], event['order']]
    row['status'] = 'Расхождение даты начала; дата таблицы сохранена'
    row['changes'] = f'Начало: таблица {row["values"][3]}; приказ {event["start"]}. Взаимно однозначное соответствие ФИО, части и длительности; требуется проверка.'
unmatched = still_unmatched
for row in rows:
    if row['event']:
        order_start = row['event'].get('sourceStart', row['event']['start'])
        if row['values'][3] != order_start:
            row['dateDifference'] = {'current': row['values'][3], 'order': order_start}
            date_note = f'Начало: таблица {row["values"][3]}; приказ {order_start}. Дата таблицы сохранена.'
            if date_note not in row['changes']:
                row['changes'] = '; '.join(filter(None, [row['changes'], date_note]))
for event in unmatched:
    end = (dt.date.fromisoformat(event['start']) + dt.timedelta(days=event['days'])).isoformat()
    rows.append({'values': [event['rank'], event['name'], event['unit'], event['start'], event['days'], end, event['basis'], event['order'], None], 'sourceRow': None, 'status': 'Добавлено', 'event': event, 'changes': ''})
# Проверенные блоки и люди сохраняют последовательность исходных приказов.
rows.sort(key=lambda r: (0, r['sourceRow']) if r['status'] == 'Вне периода проверки' else (1, r['event']['sequence']) if r['event'] else (2, r['sourceRow']))
p['rows'] = rows
p['counts'] = dict(collections.Counter(r['status'] for r in rows))
with gzip.open(OUT / 'reconciled.json.gz', 'wt', encoding='utf-8') as f:
    json.dump(p, f, ensure_ascii=False)
print(json.dumps({'counts': p['counts'], 'added_by_order': dict(collections.Counter(e['order'] for e in unmatched))}, ensure_ascii=False, indent=2))