import unittest

from keystone_model_crawler import (
    OUTPUT_COLUMNS,
    _partslink_from_text,
    parse_year_bounds,
    split_categories,
)


class ModelCrawlerTests(unittest.TestCase):
    def test_category_cell_accepts_multiple_names_and_deduplicates(self):
        self.assertEqual(split_categories("Capa,NSF,OE,Capa"), ["Capa", "NSF", "OE"])

    def test_year_bounds_use_oldest_and_newest_years(self):
        self.assertEqual(parse_year_bounds("2018 Toyota Camry fits 2019-2022"), ("2018", "2022"))

    def test_output_includes_requested_number_slots_and_description(self):
        self.assertIn("Interchange Number 5", OUTPUT_COLUMNS)
        self.assertIn("OEM Number 5", OUTPUT_COLUMNS)
        self.assertIn("Description", OUTPUT_COLUMNS)

    def test_partslink_suffix_is_preserved(self):
        self.assertEqual(_partslink_from_text("TO1000432C Available Local"), "TO1000432C")


if __name__ == "__main__":
    unittest.main()