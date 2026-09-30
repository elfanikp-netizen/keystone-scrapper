import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

import pandas as pd

from keystone_crawler import (
    MULTIPLE_NUMBER_FIELD,
    NUMBER_VALUES_FIELD,
    apply_partslink_mapping,
    extract,
    extract_partslink_number,
    is_oem_candidate,
    join_number_values,
    normalize_number_tokens,
    output_columns,
    read_partslink_mapping,
    split_number_values,
)


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

    def test_partslink_number_requires_two_letters_and_seven_digits(self):
        self.assertEqual(extract_partslink_number("Matched GM1095205"), "GM1095205")
        self.assertEqual(extract_partslink_number("GM109520"), "")
        self.assertEqual(extract_partslink_number("XGM1095205X"), "")

    def test_interchange_report_adds_partslink_column_without_number_leakage(self):
        class Page:
            def evaluate(self, script):
                return {
                    "tables": [],
                    "dl": [],
                    "body": "Partslink Number: GM1095205\nInterchange Number: 101-60187\nOEM Number: 2928851925",
                }

        with patch("keystone_crawler.extract_oem_tab_values", return_value={
            "interchange": "101-60187", "oem": "2928851925", "multiple": False,
        }):
            rows = extract(Page(), [], "101-60187", "Interchange Number")

        self.assertIn("Partslink Number", output_columns("Interchange Number"))
        self.assertEqual(output_columns("Partslink Number").count("Partslink Number"), 1)
        self.assertEqual(rows[0]["Partslink Number"], "GM1095205")
        self.assertEqual(rows[0]["Interchange Number"], "101-60187")
        self.assertEqual(rows[0]["OEM Number"], "2928851925")
        self.assertEqual(rows[0][NUMBER_VALUES_FIELD], "2928851925, 101-60187")
        self.assertFalse(rows[0][MULTIPLE_NUMBER_FIELD])

    def test_interchange_search_leaves_partslink_blank_when_result_has_none(self):
        class Page:
            def evaluate(self, script):
                return {
                    "tables": [],
                    "dl": [],
                    "body": "Unrelated inventory code: ZZ1234567",
                }

        with patch("keystone_crawler.extract_oem_tab_values", return_value={
            "interchange": "101-60187", "oem": "2928851925", "multiple": False,
        }):
            rows = [
                extract(
                    Page(),
                    [("response", {"unrelatedCode": "ZZ1234567"})],
                    "101-60187",
                    column,
                )[0]
                for column in ("Interchange Number", "Search Interchange Number")
            ]

        for row in rows:
            self.assertEqual(row["Partslink Number"], "")
            self.assertEqual(row["Interchange Number"], "101-60187")
            self.assertEqual(row["OEM Number"], "2928851925")
            self.assertEqual(row[NUMBER_VALUES_FIELD], "2928851925, 101-60187")

    def test_interchange_search_does_not_treat_oem_code_as_partslink(self):
        class Page:
            def evaluate(self, script):
                return {"tables": [], "dl": [], "body": ""}

        with patch("keystone_crawler.extract_oem_tab_values", return_value={
            "interchange": "101-60187", "oem": "2928851925", "multiple": False,
        }):
            rows = extract(
                Page(),
                [("response", {"inventoryIdentifier": "ZZ1234567", "brand": "Example"})],
                "101-60187",
                "Interchange Number",
            )

        self.assertEqual(rows[0]["Partslink Number"], "")
        self.assertEqual(rows[0]["OEM Number"], "2928851925")
        self.assertEqual(rows[0]["Interchange Number"], "101-60187")

    def test_interchange_search_uses_only_unambiguous_source_pair(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            pd.DataFrame({
                "Partslink Number": ["HY1115122", "HO1115107", "KI1115122"],
                "Interchange Number": ["101-60187", "101-MB1D26", "101-MB1D26"],
            }).to_excel(input_path, index=False)

            mapping = read_partslink_mapping(input_path, "Interchange Number")

        self.assertEqual(mapping, {"101-60187": "HY1115122"})

    def test_source_partslink_pair_replaces_catalog_card_identifier(self):
        row = {"Partslink Number": "MB1000490"}
        apply_partslink_mapping(row, "101-60187", {"101-60187": "HY1115122"})
        self.assertEqual(row["Partslink Number"], "HY1115122")

    def test_missing_source_pair_keeps_partslink_blank(self):
        row = {"Partslink Number": ""}
        apply_partslink_mapping(row, "101-60187", {})
        self.assertEqual(row["Partslink Number"], "")

    def test_interchange_search_value_populates_interchange_and_number_values(self):
        class Page:
            def evaluate(self, script):
                return {"tables": [], "dl": [], "body": "Partslink Number: GM1095205"}

        with patch("keystone_crawler.extract_oem_tab_values", return_value={
            "interchange": "", "oem": "2928851925", "multiple": False,
        }) as extract_values:
            rows = extract(Page(), [], "101-60187", "Interchange Number")

        extract_values.assert_called_once_with(
            unittest.mock.ANY,
            "101-60187",
            exclude_search_value=False,
        )
        self.assertEqual(rows[0]["Partslink Number"], "GM1095205")
        self.assertEqual(rows[0]["Interchange Number"], "101-60187")
        self.assertEqual(rows[0]["Interchange 1"], "101-60187")
        self.assertEqual(rows[0]["OEM Number"], "2928851925")
        self.assertEqual(rows[0][NUMBER_VALUES_FIELD], "2928851925, 101-60187")

    def test_partslink_code_is_excluded_from_number_fields_only(self):
        parsed = normalize_number_tokens("GM1095205, 101-60187, 2928851925")
        self.assertEqual(parsed, ["101-60187", "2928851925"])
        self.assertEqual(
            join_number_values("GM1095205, 101-60187", "GM1095205, 2928851925"),
            "101-60187, 2928851925",
        )


if __name__ == "__main__":
    unittest.main()
