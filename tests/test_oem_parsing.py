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
    extract_fitments_tab_values,
    extract_partslink_number,
    is_oem_candidate,
    join_number_values,
    normalize_number_tokens,
    extract_oem_tab_values,
    load_results,
    output_columns,
    parse_fitments_text,
    read_partslink_mapping,
    read_parts,
    save_results,
    split_number_values,
)


class OemParsingTests(unittest.TestCase):
    def test_input_parts_keep_duplicate_rows(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "parts.xlsx"
            pd.DataFrame({"Partslink Number": ["GM1240366V", "GM1240366V"]}).to_excel(
                input_path, index=False
            )

            self.assertEqual(read_parts(input_path, "Partslink Number"), [
                "GM1240366V", "GM1240366V",
            ])

    def test_saved_results_repeat_once_per_duplicate_input_row(self):
        with TemporaryDirectory() as directory:
            output_path = Path(directory) / "results.xlsx"
            query = "GM1240366V"
            result = {column: "" for column in output_columns()}
            result.update({"Partslink Number": query, "Status": "OK"})

            save_results(output_path, [query, query], {query: [result]})
            saved = pd.read_excel(output_path, dtype=str).fillna("")

        self.assertEqual(saved["Partslink Number"].tolist(), [query, query])

    def test_existing_single_result_expands_to_duplicate_input_rows(self):
        with TemporaryDirectory() as directory:
            output_path = Path(directory) / "results.xlsx"
            query = "GM1240366V"
            result = {column: "" for column in output_columns()}
            result.update({"Partslink Number": query, "Status": "OK"})
            save_results(output_path, [query], {query: [result]})

            results = load_results(output_path)
            save_results(output_path, [query, query], results)
            saved = pd.read_excel(output_path, dtype=str).fillna("")

        self.assertEqual(saved["Partslink Number"].tolist(), [query, query])

    def test_fitments_text_produces_year_bounds_brand_and_model(self):
        text = """2017 GMC TERRAIN
2015 GMC TERRAIN
2013 GMC TERRAIN
2011 GMC TERRAIN
2016 GMC TERRAIN
2014 GMC TERRAIN
2012 GMC TERRAIN
2010 GMC TERRAIN"""

        self.assertEqual(parse_fitments_text(text), {
            "Oldest Year": "2010",
            "Newest Year": "2017",
            "Brand": "GMC",
            "Model": "TERRAIN",
        })

    def test_fitments_text_replaces_underscores_in_vehicle_fields(self):
        self.assertEqual(parse_fitments_text("2018_GMC_TERRAIN_SLE"), {
            "Oldest Year": "2018",
            "Newest Year": "2018",
            "Brand": "GMC",
            "Model": "TERRAIN SLE",
        })

    def test_fitments_are_read_from_active_fitments_tab(self):
        class Tab:
            first = None

            def __init__(self):
                self.first = self
                self.clicked = False

            def is_visible(self):
                return True

            def click(self, **kwargs):
                self.clicked = True

            def count(self):
                return 1

            def inner_text(self):
                return "2017 GMC TERRAIN\n2010 GMC TERRAIN"

        class Page:
            def __init__(self):
                self.tab = Tab()

            def locator(self, selector):
                if selector == 'button:has-text("Fitments")':
                    return self.tab
                if selector == '.mat-tab-body-wrapper .mat-tab-body-active':
                    return self.tab
                return TabCollection()

        class TabCollection:
            first = None

            def __init__(self):
                self.first = self

            def is_visible(self):
                return False

            def count(self):
                return 0

        page = Page()
        with patch("keystone_crawler.wait_settled"):
            values = extract_fitments_tab_values(page)

        self.assertTrue(page.tab.clicked)
        self.assertEqual(values, {
            "Oldest Year": "2010",
            "Newest Year": "2017",
            "Brand": "GMC",
            "Model": "TERRAIN",
        })

    def test_fitments_values_override_other_extracted_vehicle_fields(self):
        class Page:
            def evaluate(self, script):
                return {
                    "tables": [],
                    "dl": [],
                    "body": "Oldest Year: 2008\nNewest Year: 2020\nBrand: Other\nModel: Vehicle",
                }

        fitments = {
            "Oldest Year": "2010",
            "Newest Year": "2017",
            "Brand": "GMC",
            "Model": "TERRAIN",
        }
        with patch("keystone_crawler.extract_fitments_tab_values", return_value=fitments), \
                patch("keystone_crawler.extract_oem_tab_values", return_value={}):
            rows = extract(Page(), [], "GM1240366V")

        self.assertEqual(
            {field: rows[0][field] for field in fitments},
            fitments,
        )

    def test_fitments_alone_still_produce_a_result_row(self):
        class Page:
            def evaluate(self, script):
                return {"tables": [], "dl": [], "body": ""}

        fitments = {
            "Oldest Year": "2010",
            "Newest Year": "2017",
            "Brand": "GMC",
            "Model": "TERRAIN",
        }
        with patch("keystone_crawler.extract_fitments_tab_values", return_value=fitments), \
                patch("keystone_crawler.extract_oem_tab_values", return_value={}):
            rows = extract(Page(), [], "GM1240366V")

        self.assertEqual(len(rows), 1)
        self.assertEqual(
            {field: rows[0][field] for field in fitments},
            fitments,
        )

    def test_numeric_oem_values_are_accepted(self):
        self.assertTrue(is_oem_candidate("22818031"))
        self.assertTrue(is_oem_candidate("16688518389999"))
        self.assertTrue(is_oem_candidate("1668851838649999"))
        self.assertTrue(is_oem_candidate("A12345"))
        self.assertTrue(is_oem_candidate("12345A"))

    def test_long_numeric_oem_values_are_parsed_from_interchange_result_tab(self):
        class EmptyLocator:
            first = None

            def count(self):
                return 0

            def is_visible(self):
                return False

        class TabBody:
            def inner_text(self):
                return "\n".join([
                    "101-10026",
                    "HY1115122",
                    "1668851838649999",
                    "16688518389999",
                    "MB1000528",
                ])

        class BodyLocator:
            def count(self):
                return 1

            def nth(self, index):
                return TabBody()

        class Page:
            def locator(self, selector):
                if selector == ".mat-tab-body-wrapper .mat-tab-body":
                    return BodyLocator()
                if selector == ".mat-tab-body-wrapper .mat-tab-body-active":
                    return EmptyLocator()
                return EmptyLocator()

        values = extract_oem_tab_values(Page(), "101-10026", exclude_search_value=False)
        self.assertEqual(values["interchange"], "101-10026")
        self.assertEqual(values["partslink"], "HY1115122")
        self.assertEqual(values["oem"], "1668851838649999, 16688518389999")
        self.assertEqual(
            values["number_values"],
            "101-10026, 1668851838649999, 16688518389999",
        )

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

    def test_interchange_search_prefers_partslink_number_from_oem_tab(self):
        class Page:
            def evaluate(self, script):
                return {
                    "tables": [],
                    "dl": [],
                    "body": "Partslink Number: GM1095205",
                }

        with patch("keystone_crawler.extract_oem_tab_values", return_value={
            "interchange": "101-60187",
            "oem": "2928851925",
            "partslink": "HY1115122",
            "multiple": False,
        }):
            rows = extract(Page(), [], "101-60187", "Interchange Number")

        self.assertEqual(rows[0]["Partslink Number"], "HY1115122")

    def test_oem_tab_partslink_overrides_catalog_record_for_interchange_search(self):
        class Page:
            def evaluate(self, script):
                return {"tables": [], "dl": [], "body": ""}

        catalog_record = {
            "Partslink Number": "MB1000490",
            "Interchange Number": "101-60187",
            "OEM Number": "2928851925",
        }
        with patch(
            "keystone_crawler.walk_json",
            side_effect=lambda body, path, records: records.append(catalog_record),
        ), patch("keystone_crawler.extract_oem_tab_values", return_value={
            "interchange": "101-60187",
            "oem": "2928851925",
            "partslink": "HY1115122",
            "multiple": False,
        }):
            rows = extract(Page(), [("response", {})], "101-60187", "Interchange Number")

        self.assertEqual(rows[0]["Partslink Number"], "HY1115122")

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
