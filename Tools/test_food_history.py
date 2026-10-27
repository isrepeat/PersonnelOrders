"""Проверки границ периодов и противоречий Food на независимых примерах."""

import unittest
from datetime import date

from food_history import audit_person, similar


class HistoryTests(unittest.TestCase):
    def setUp(self):
        self.person = dict(name='ТЕСТЕНКО Іван Петрович', source_name='ТЕСТЕНКО Іван Петрович', rank='солдат', source_rank='солдат')

    def event(self, kind, start, end, name=None):
        return dict(kind=kind, name=name or self.person['name'], military='А5055', start=date.fromisoformat(start),
                    end=date.fromisoformat(end) if end else None, ref='Сухпрод!C17' if kind == 'dry' else 'Продовольче!C8', issue=[], ipn='')

    def group(self, kind, start, **extra):
        return dict(kind=kind, military='А5055', start=start, people=[self.person], incoming_number='1', **extra)

    def test_adjacent_dry_is_allowed(self):
        result = audit_person(self.group('dry', '2026-10-24', duration=3), self.person, [self.event('dry', '2026-10-21', '2026-10-24')])
        self.assertFalse(result['conflicts'])

    def test_overlapping_dry_is_conflict(self):
        result = audit_person(self.group('dry', '2026-10-24', duration=3), self.person, [self.event('dry', '2026-10-23', '2026-10-25')])
        self.assertTrue(any('Перетин сухпайка' in x for x in result['conflicts']))

    def test_closed_enrollment_cannot_be_removed_again(self):
        result = audit_person(self.group('removal', '2026-10-24'), self.person, [self.event('food', '2026-10-01', '2026-10-20')])
        self.assertTrue(result['conflicts'])

    def test_removal_of_active_person(self):
        result = audit_person(self.group('removal', '2026-10-24'), self.person, [self.event('food', '2026-10-01', None)])
        self.assertFalse(result['conflicts'])

    def test_enrollment_during_dry(self):
        result = audit_person(self.group('enrollment', '2026-10-24'), self.person, [self.event('dry', '2026-10-23', '2026-10-26')])
        self.assertTrue(result['conflicts'])

    def test_dry_ceases_on_enrollment_day(self):
        result = audit_person(self.group('enrollment', '2026-10-26'), self.person, [self.event('dry', '2026-10-23', '2026-10-26')])
        self.assertFalse(result['conflicts'])

    def test_duplicate_enrollment(self):
        result = audit_person(self.group('enrollment', '2026-10-24'), self.person, [self.event('food', '2026-10-01', None)])
        self.assertTrue(result['conflicts'])

    def test_alias_is_found_even_with_exact_history(self):
        result = audit_person(self.group('dry', '2026-10-24', duration=3), self.person,
                              [self.event('dry', '2026-10-01', '2026-10-04'), self.event('dry', '2026-10-24', '2026-10-25', 'ТЕСТЕНК Іван Петрович')])
        self.assertTrue(result['history_alias'])
        self.assertTrue(any('Можливе протиріччя' in x for x in result['warnings']))

    def test_surname_alone_is_not_identity(self):
        self.assertFalse(similar('ТЕСТЕНКО Олександр Васильович', {'ТЕСТЕНКО Іван Петрович'}))

    def test_current_package_removal_after_enrollment(self):
        first = self.group('enrollment', '2026-10-20')
        second = self.group('removal', '2026-10-24')
        result = audit_person(second, self.person, [], [first, second])
        self.assertFalse(result['conflicts'])

    def test_opposite_current_instructions_same_day(self):
        first = self.group('enrollment', '2026-10-24')
        second = self.group('removal', '2026-10-24')
        result = audit_person(first, self.person, [], [first, second])
        self.assertTrue(result['conflicts'])

    def test_existing_same_report_is_not_new_overlap(self):
        group = self.group('dry', '2026-10-24', duration=3)
        event = self.event('dry', '2026-10-24', '2026-10-27')
        event['basis'] = 'рапорт (вх. № 1 від 23.10.2026)'
        result = audit_person(group, self.person, [event])
        self.assertFalse(result['conflicts'])
        self.assertIn('уже внесений', result['notes'])

    def test_close_candidate_with_open_food_is_uncertain_not_definite_absence(self):
        result = audit_person(self.group('dry', '2026-10-24', duration=3), self.person,
                              [self.event('food', '2026-06-01', '2026-07-01'), self.event('food', '2026-07-20', None, 'ТЕСТЕНК Іван Петрович')])
        self.assertFalse(result['conflicts'])
        self.assertTrue(result['history_alias'])
        self.assertTrue(any('покриває сухпай' in text for text in result['warnings']))


if __name__ == '__main__':
    unittest.main()