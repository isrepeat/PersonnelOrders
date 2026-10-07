"""Проверки извлечения ТВО без персональных данных из реального справочника."""
import unittest

from extract_order import acting_replacement, return_before_departure


class ActingReplacementTests(unittest.TestCase):
    def setUp(self):
        self.alf = {"1234567890": "ТЕСТЕНКО Олександр Юрійович"}
        self.positions = {"1234567890": "A1A123"}
        self.assignment = "Тимчасове виконання обов'язків командира покласти на капітана ТЕСТЕНКА Олександра Юрійовича."

    def test_full_name_and_position(self):
        result, warning = acting_replacement(self.assignment, self.alf, self.positions)
        self.assertEqual(result, {"ТВО.ПІБ": self.alf["1234567890"], "ТВО.ІПН": "1234567890", "ТВО.Посада": "РОЗП"})
        self.assertFalse(warning)

    def test_return_is_not_assignment(self):
        result, warning = acting_replacement("Капітану ТЕСТЕНКУ повернутись до виконання обов'язків.", self.alf, self.positions)
        self.assertFalse(any(result.values()))
        self.assertFalse(warning)

    def test_ambiguous_name_is_not_first_match(self):
        self.alf["0987654321"] = self.alf["1234567890"]
        result, warning = acting_replacement(self.assignment, self.alf, self.positions)
        self.assertFalse(any(result.values()))
        self.assertTrue(warning)

    def test_explicit_ipn_resolves_duplicate(self):
        self.alf["0987654321"] = self.alf["1234567890"]
        result, warning = acting_replacement(self.assignment.replace("Юрійовича.", "Юрійовича, 1234567890."), self.alf, self.positions)
        self.assertEqual(result["ТВО.ІПН"], "1234567890")
        self.assertFalse(warning)

    def test_different_given_name_does_not_match(self):
        result, warning = acting_replacement(self.assignment.replace("Олександра", "Володимира"), self.alf, self.positions)
        self.assertFalse(any(result.values()))
        self.assertTrue(warning)


class ReturnBeforeDepartureTests(unittest.TestCase):
    def row(self, departure="01 жовтня 2026"):
        return {"Лист": "Відсутні", "Дія": "Вибув у відрядження", "Наказ": "№284 від 01.10.2026",
                "Дата події": "01.10.2026", "Продовольча дата": "НЕ ЗМІНЮВАТИ", "Супровідний документ": "2026/284/13",
                "Текст наказу": "Самовільно залишив частину. Вважати таким, що з 29 вересня 2026 року повернувся до військової частини А7018. "
                f"Направити у відрядження у військову частину А0501 з {departure} року."}

    def test_different_dates_have_bridge(self):
        row = self.row()
        closed, bridge = return_before_departure(row)
        self.assertEqual(closed["Прибуття"], "29.09.2026")
        self.assertEqual((bridge["Вибуття"], bridge["Прибуття"], bridge["Куди"]), ("29.09.2026", "01.10.2026", "А7018"))
        self.assertEqual((row["Дата події"], row["Куди"]), ("01.10.2026", "А0501"))
        self.assertEqual(bridge["Супровідний документ"], "")
        self.assertNotIn("Прибуття.Продовольче", bridge)

    def test_same_dates_no_bridge(self):
        self.assertEqual(len(return_before_departure(self.row("29 вересня 2026"))), 1)

    def test_reverse_dates_flagged(self):
        row = self.row("28 вересня 2026")
        self.assertEqual(return_before_departure(row), [])
        self.assertIn("Проверить", row["Деталі"])


if __name__ == "__main__":
    unittest.main()