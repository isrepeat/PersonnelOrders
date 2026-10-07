import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import compare_excel


class CompareExcelTests(unittest.TestCase):
    def generate(self, old_rows, new_rows):
        args = SimpleNamespace(old=Path('old.xlsx'), new=Path('new.xlsx'), key='ПІБ', threshold=0.65)
        with patch.object(compare_excel, 'read', side_effect=[{'Лист': old_rows}, {'Лист': new_rows}]):
            return compare_excel.generate(args)

    def test_reordered_duplicates_and_added_deleted_records(self):
        header = (1, ('ПІБ', 'Код', 'Дата', 'Примітка'))
        alice = ('Анна', '1', '01.10.2026', '')
        bob = ('Богдан', '2', '01.10.2026', '')
        charlie = ('Василь', '3', '02.10.2026', '')
        report, totals = self.generate(
            [header, (2, alice), (3, alice), (4, bob)],
            [header, (2, alice), (3, charlie), (4, alice)],
        )
        self.assertEqual((totals['change'], totals['add'], totals['delete']), (0, 1, 1))
        self.assertEqual(report.count('<article '), 2)

    def test_replacements_empty_values_and_escaped_html(self):
        header = (1, ('ПІБ', 'Код', 'Дата', 'Примітка', 'Поле'))
        report, totals = self.generate(
            [header, (2, ('Анна', '1', '01.10.2026', 'старое', ''))],
            [header, (2, ('Анна', '1', '01.10.2026', '<script>новое</script>', ''))],
        )
        self.assertEqual(totals['change'], 1)
        self.assertIn('class="cell replacement"', report)
        self.assertIn('Было · D2', report)
        self.assertIn('&lt;script&gt;новое&lt;/script&gt;', report)
        self.assertNotIn('<script>новое</script>', report)

    def test_added_value_is_green_and_formula_is_collapsed(self):
        header = (1, ('ПІБ', 'Код', 'Дата', 'Примітка'))
        report, totals = self.generate(
            [header, (2, ('Анна', '1', '01.10.2026', ''))],
            [header, (2, ('Анна', '1', '01.10.2026', '=SUM(A1:A2)'))],
        )
        self.assertEqual(totals['change'], 1)
        self.assertIn('class="cell insertion"', report)
        self.assertIn('<details class="formula">', report)

    def test_all_named_columns_are_preserved(self):
        header = (1, ('', 'ПІБ', 'Код', 'Дата', 'Примітка', 'Пустое'))
        report, totals = self.generate(
            [header, (2, ('', 'Анна', '1', '01.10.2026', 'старое', ''))],
            [header, (2, ('', 'Анна', '1', '01.10.2026', 'новое', ''))],
        )
        self.assertEqual(totals['change'], 1)
        self.assertEqual(report.count('class="cell '), 5)
        self.assertEqual(report.count('<pre>-- // --</pre>'), 4)
        self.assertLess(report.index('<b>ПІБ</b>'), report.index('<b>Код</b>'))
        self.assertLess(report.index('<b>Код</b>'), report.index('<b>Дата</b>'))
        self.assertIn('<b>Пустое</b>', report)

    def test_changed_headers_stop_comparison(self):
        with self.assertRaisesRegex(RuntimeError, 'Изменены заголовки'):
            self.generate([(1, ('ПІБ', 'Код'))], [(1, ('ПІБ', 'Дата'))])

    def test_food_and_dry_rations_are_available_without_other_tabs(self):
        args = SimpleNamespace(old=Path('old.xlsx'), new=Path('new.xlsx'), key='ПІБ', threshold=0.65)
        header = (2, ('Звання', 'ПІБ', 'Початок', 'Тривалість', 'Припинення'))
        old = {'Продовольче': [header], 'Сухпрод': [header]}
        new = {'Продовольче': [header], 'Сухпрод': [header, (3, ('солдат', 'Анна', '06.10.2026', '3', '09.10.2026'))]}
        with patch.object(compare_excel, 'read', side_effect=[old, new]):
            report, totals = compare_excel.generate(args)
        self.assertEqual(totals['add'], 1)
        self.assertIn("['Продовольче','Продовольче']", report)
        self.assertIn("['Сухпрод','Сухпрод']", report)
        self.assertIn('06.10.2026', report)
        self.assertIn("button.textContent===pageNames[index][1]", report)

    def test_equal_candidates_are_flagged(self):
        header = (1, ('ПІБ', 'Код', 'Дата', 'Примітка'))
        report, totals = self.generate(
            [header, (2, ('Анна', '1', '01.10.2026', 'старое'))],
            [header, (2, ('Анна', '1', '01.10.2026', 'первое')), (3, ('Анна', '1', '01.10.2026', 'второе'))],
        )
        self.assertEqual((totals['change'], totals['add']), (1, 1))
        self.assertIn('Несколько записей имеют одинаковую близость', report)


if __name__ == '__main__':
    unittest.main()