import unittest

from app import calculate_risk


class CoreTest(unittest.TestCase):
    def test_risk_classification_low_medium_high(self):
        self.assertEqual(calculate_risk(1, 1), (1, "Baixo"))
        self.assertEqual(calculate_risk(2, 2), (4, "Médio"))
        self.assertEqual(calculate_risk(3, 3), (9, "Alto"))


if __name__ == "__main__":
    unittest.main()
