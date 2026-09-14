import tempfile
import unittest
from pathlib import Path

from database import Database


class DatabaseTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.db = Database(Path(self.tempdir.name) / "test.db")
        self.period = "2026-09"
        self.a = self.db.add_student("Perez", "Ana", "1")
        self.b = self.db.add_student("Gomez", "Bruno", "2")

    def tearDown(self):
        self.tempdir.cleanup()

    def test_initial_balance_is_100(self):
        self.assertEqual(self.db.get_balance(self.a, self.period).available, 100)

    def test_print_reduces_balance(self):
        self.db.register_print(self.a, 12, self.period, "apunte.pdf")
        balance = self.db.get_balance(self.a, self.period)
        self.assertEqual(balance.printed, 12)
        self.assertEqual(balance.available, 88)

    def test_donation_moves_quota_between_students(self):
        self.db.donate(self.a, self.b, 25, self.period)
        self.assertEqual(self.db.get_balance(self.a, self.period).available, 75)
        self.assertEqual(self.db.get_balance(self.b, self.period).available, 125)

    def test_cannot_donate_more_than_available(self):
        with self.assertRaises(ValueError):
            self.db.donate(self.a, self.b, 101, self.period)

    def test_reversal_reintegrates_last_print(self):
        self.db.register_print(self.a, 20, self.period, "fallo.pdf")
        self.db.reverse_last_print(self.a, self.period, "Impresora trabada")
        balance = self.db.get_balance(self.a, self.period)
        self.assertEqual(balance.printed, 0)
        self.assertEqual(balance.corrected, 20)
        self.assertEqual(balance.available, 100)

    def test_same_print_cannot_be_reversed_twice(self):
        self.db.register_print(self.a, 10, self.period, "x.pdf")
        self.db.reverse_last_print(self.a, self.period, "Error")
        with self.assertRaises(ValueError):
            self.db.reverse_last_print(self.a, self.period, "Otra vez")

    def test_periods_are_independent(self):
        self.db.register_print(self.a, 30, "2026-08")
        self.assertEqual(self.db.get_balance(self.a, "2026-08").available, 70)
        self.assertEqual(self.db.get_balance(self.a, "2026-09").available, 100)


if __name__ == "__main__":
    unittest.main()
