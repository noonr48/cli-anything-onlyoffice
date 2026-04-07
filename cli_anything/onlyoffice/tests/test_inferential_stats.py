import os
import tempfile
import unittest

from cli_anything.onlyoffice.utils.docserver import get_client


class OnlyOfficeInferentialStatsTests(unittest.TestCase):
    def setUp(self):
        self.client = get_client()
        self.tmpdir = tempfile.TemporaryDirectory(prefix="oo-inferential-")
        self.path = os.path.join(self.tmpdir.name, "stats.xlsx")

        # A=Group, B=Employability score, C=Preparedness score, D=Needs experience
        self.client.write_spreadsheet(
            output_path=self.path,
            headers=["Group", "Score1", "Score2", "NeedsExp"],
            data=[
                [1, 10, 20, 1],
                [1, 11, 22, 1],
                [1, 9, 18, 1],
                [1, 10, 20, 1],
                [2, 20, 40, 2],
                [2, 21, 42, 2],
                [2, 19, 38, 2],
                [2, 20, 40, 2],
            ],
            sheet_name="Sheet0",
            overwrite_workbook=True,
        )

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_frequency_table(self):
        result = self.client.frequencies(self.path, "A", sheet_name="Sheet0")
        self.assertTrue(result["success"])
        self.assertEqual(result["valid_n"], 8)
        freq = {row["category"]: row["count"] for row in result["frequencies"]}
        self.assertEqual(freq[1], 4)
        self.assertEqual(freq[2], 4)

    def test_frequency_table_with_valid_filter(self):
        # Add a junk category row and ensure filtering excludes it
        self.client.append_to_spreadsheet(
            self.path,
            ["unknown", "", "", ""],
            sheet_name="Sheet0",
        )
        result = self.client.frequencies(
            self.path,
            "A",
            sheet_name="Sheet0",
            allowed_values=["1", "2"],
        )
        self.assertTrue(result["success"])
        self.assertEqual(result["valid_n"], 8)
        self.assertEqual(result["excluded_n"], 1)

    def test_correlation(self):
        result = self.client.correlation_test(
            self.path, "B", "C", sheet_name="Sheet0", method="pearson"
        )
        self.assertTrue(result["success"])
        self.assertAlmostEqual(result["statistic"], 1.0, places=6)
        self.assertLess(result["p_value"], 1e-5)

    def test_ttest(self):
        result = self.client.ttest_independent(
            self.path,
            value_column="B",
            group_column="A",
            group_a="1",
            group_b="2",
            sheet_name="Sheet0",
            equal_var=False,
        )
        self.assertTrue(result["success"])
        self.assertLess(result["p_value"], 1e-5)
        self.assertAlmostEqual(result["mean_a"], 10.0, places=6)
        self.assertAlmostEqual(result["mean_b"], 20.0, places=6)

    def test_chi_square(self):
        result = self.client.chi_square_test(self.path, "A", "D", sheet_name="Sheet0")
        self.assertTrue(result["success"])
        self.assertLess(result["p_value"], 0.05)
        self.assertGreater(result["cramers_v"], 0.7)

    def test_chi_square_with_valid_filters(self):
        self.client.append_to_spreadsheet(
            self.path,
            ["junk", 12, 24, "other"],
            sheet_name="Sheet0",
        )
        result = self.client.chi_square_test(
            self.path,
            "A",
            "D",
            sheet_name="Sheet0",
            row_allowed_values=["1", "2"],
            col_allowed_values=["1", "2"],
        )
        self.assertTrue(result["success"])
        self.assertEqual(result["excluded_n"], 1)
        self.assertEqual(result["rows"], [1, 2])
        self.assertEqual(result["cols"], [1, 2])


if __name__ == "__main__":
    unittest.main()
