import tempfile
import unittest
from pathlib import Path

import pandas as pd

import keystone_model_crawler as crawler


class KeystoneModelResumeTests(unittest.TestCase):
    def test_saved_input_key_skips_success_and_no_data_but_not_errors(self) -> None:
        first = {
            "Year": "2020",
            "Brand": "Ford",
            "Model": "Escape",
            "Parts": "Front Bumper",
            "Category": "Body",
        }
        duplicate = {**first, "Brand": " ford "}
        other = {**first, "Category": "Lighting"}
        rows = [
            {"Search Key": crawler.search_key(first), "Search Status": "OK"},
            {"Search Key": crawler.search_key(other), "Search Status": "NO DATA"},
            {"Search Key": crawler.search_key({**first, "Parts": "Mirror"}), "Search Status": "ERROR"},
        ]

        completed = crawler.completed_search_keys(rows)

        self.assertEqual(crawler.search_key(first), crawler.search_key(duplicate))
        self.assertIn(crawler.search_key(first), completed)
        self.assertIn(crawler.search_key(other), completed)
        self.assertNotIn(crawler.search_key({**first, "Parts": "Mirror"}), completed)

    def test_checkpoint_workbook_round_trips_search_metadata(self) -> None:
        record = {
            "Year": "2020",
            "Brand": "Ford",
            "Model": "Escape",
            "Parts": "Front Bumper",
            "Category": "Body",
        }
        output_row = {
            "Partslink Number": "FO1234567",
            "Brand": "Ford",
            "Model": "Escape",
            "Search Key": crawler.search_key(record),
            "Search Status": "OK",
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "results.xlsx"
            crawler.save_results(output_path, [output_row])
            saved_rows = pd.read_excel(output_path, dtype=str).fillna("").to_dict("records")

        self.assertIn(crawler.search_key(record), crawler.completed_search_keys(saved_rows))


if __name__ == "__main__":
    unittest.main()