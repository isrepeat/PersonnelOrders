"""Сохраняет первую версию появления записи при пересборке таблицы."""

import re
from pathlib import Path

from openpyxl import load_workbook

from food_history import day, normal


def record_keys(kind, military, start, duration, source, names, ipn):
    source_files = tuple(sorted(Path(value.strip()).name for value in str(source or '').split(';') if value.strip()))
    prefix = (kind, normal(military), day(start).isoformat(), int(duration or 0), source_files)
    identities = {normal(name) for name in names if name}
    if ipn:
        identities.add('ІПН:' + str(ipn).strip())
    return [prefix + (identity,) for identity in identities]


def apply_versions(run, output):
    output = Path(output).resolve()
    match = re.match(r'v(\d+) ', output.name)
    if not match:
        raise ValueError('Имя таблицы должно начинаться с vN')
    version = int(match[1])
    known = {}
    files = []
    for path in output.parent.glob('v*.xlsx'):
        found = re.match(r'v(\d+) ', path.name)
        if found and int(found[1]) < version:
            files.append((int(found[1]), path))
    for number, path in sorted(files):
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            for sheet in book:
                if sheet.title not in ('Продовольче', 'Сухпрод', 'Проблеми'):
                    continue
                rows = sheet.iter_rows(min_row=2, values_only=True)
                headers = {value: index for index, value in enumerate(next(rows)) if value}
                def get(row, name):
                    return row[headers[name]] if name in headers else None
                for row in rows:
                    name = get(row, 'ПІБ')
                    if not name:
                        continue
                    if sheet.title == 'Продовольче':
                        enrolled = get(row, 'Зарахування')
                        kind, start = ('enrollment', enrolled) if enrolled else ('removal', get(row, 'Виключення'))
                        duration = 0
                    else:
                        kind = 'dry'
                        start = get(row, 'Початок') or get(row, 'Запитаний початок')
                        duration = get(row, 'Тривалість') or get(row, 'Запитана тривалість')
                    if not start:
                        continue
                    keys = record_keys(kind, get(row, 'Установа') or get(row, 'ІПН/Установа'), start, duration,
                                       get(row, 'Джерело'), [name], get(row, 'ІПН'))
                    first = str(get(row, 'Додано у версії') or f'v{number}')
                    for key in keys:
                        known.setdefault(key, first)
        finally:
            book.close()
    def assign(group, person):
        keys = record_keys(group['kind'], group['military'], group['start'], group.get('duration'),
                           '; '.join(group['source_pages']), [person['name'], person.get('source_name')], person.get('ipn'))
        found = [known[key] for key in keys if key in known]
        person['added_version'] = min(found, key=lambda value: int(value[1:])) if found else f'v{version}'
    for group in run['groups']:
        for person in group['people']:
            assign(group, person)
    for item in run['problems']:
        assign(item['group'], item['person'])
    return run