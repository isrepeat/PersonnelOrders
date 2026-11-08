"""Проверка сохранения версии при перестановке и переносе записей."""

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from food_versions import apply_versions


class VersionTests(unittest.TestCase):
    def test_existing_record_and_new_operation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            book = Workbook()
            sheet = book.active
            sheet.title = 'Сухпрод'
            sheet.append([])
            sheet.append(['ПІБ', 'ІПН/Установа', 'Початок', 'Тривалість', 'Джерело', 'ІПН'])
            sheet.append(['ТЕСТЕНКО Іван Петрович', 'А7383', datetime(2026, 11, 7), 3, 'report.jpeg', '123'])
            book.save(root / 'v1 test.xlsx')
            people = [dict(name='НОВЕНКО Іван Петрович', ipn='456'),
                      dict(name='ВИПРАВЛЕНКО Іван Петрович', ipn='123')]
            group = dict(kind='dry', military='А7383', start='2026-11-07', duration=3,
                         source_pages=['report.jpeg'], people=people)
            new_operation = dict(group, start='2026-11-10', people=[dict(people[1])])
            held = dict(group={key: value for key, value in group.items() if key != 'people'}, person=dict(people[1]))
            run = dict(groups=[group, new_operation], problems=[held])
            apply_versions(run, root / 'v2 test.xlsx')
            self.assertEqual([person['added_version'] for person in people], ['v2', 'v1'])
            self.assertEqual(new_operation['people'][0]['added_version'], 'v2')
            self.assertEqual(held['person']['added_version'], 'v1')


if __name__ == '__main__':
    unittest.main()