import unittest

from keystone_crawler import is_oem_candidate, split_number_values


class OemParsingTests(unittest.TestCase):
    def test_numeric_oem_values_are_accepted(self):
        self.assertTrue(is_oem_candidate("22818031"))
        self.assertTrue(is_oem_candidate("A12345"))
        self.assertTrue(is_oem_candidate("12345A"))

    def test_mixed_prefix_and_suffix_oem_values_are_accepted(self):
        self.assertTrue(is_oem_candidate("KT4216138A"))
        self.assertTrue(is_oem_candidate("KT4716138C"))

    def test_number_values_are_split_into_up_to_five_slots(self):
        self.assertEqual(
            split_number_values("A12345, B23456, C34567, D45678, E56789, F67890"),
            ["A12345", "B23456", "C34567", "D45678", "E56789"],
        )


if __name__ == "__main__":
    unittest.main()
