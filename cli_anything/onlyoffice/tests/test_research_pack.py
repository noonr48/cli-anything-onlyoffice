import os
import tempfile
import unittest

from cli_anything.onlyoffice.utils.docserver import get_client


class OnlyOfficeResearchPackTests(unittest.TestCase):
    def setUp(self):
        self.client = get_client()
        self.tmp = tempfile.TemporaryDirectory(prefix="oo-pack-")
        self.path = os.path.join(self.tmp.name, "hlth.xlsx")

        # Build minimal HLTH-like sheet with columns up to AL (38)
        headers = [f"C{i}" for i in range(1, 39)]
        rows = [
            # A, F, M, R, Y, AL positions: 1,6,13,18,25,38
            [
                1,
                20,
                1,
                1,
                1,
                1,
                0,
                0,
                0,
                0,
                0,
                0,
                4,
                0,
                0,
                0,
                0,
                3,
                0,
                0,
                0,
                0,
                0,
                0,
                4,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                5,
            ],
            [
                2,
                22,
                1,
                1,
                1,
                4,
                0,
                0,
                0,
                0,
                0,
                0,
                5,
                0,
                0,
                0,
                0,
                5,
                0,
                0,
                0,
                0,
                0,
                0,
                5,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                4,
            ],
            [
                1,
                19,
                1,
                1,
                1,
                1,
                0,
                0,
                0,
                0,
                0,
                0,
                3,
                0,
                0,
                0,
                0,
                3,
                0,
                0,
                0,
                0,
                0,
                0,
                3,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                4,
            ],
            [
                2,
                21,
                1,
                1,
                1,
                4,
                0,
                0,
                0,
                0,
                0,
                0,
                5,
                0,
                0,
                0,
                0,
                4,
                0,
                0,
                0,
                0,
                0,
                0,
                5,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                3,
            ],
            # junk row/category should be filtered by profile rules
            [
                "Which sex/gender...",
                "",
                "",
                "",
                "",
                1,
                0,
                0,
                0,
                0,
                0,
                0,
                4,
                0,
                0,
                0,
                0,
                4,
                0,
                0,
                0,
                0,
                0,
                0,
                4,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                "bad",
            ],
        ]

        self.client.write_spreadsheet(
            output_path=self.path,
            headers=headers,
            data=rows,
            sheet_name="Sheet0",
            overwrite_workbook=True,
            coerce_rows=True,
        )

    def tearDown(self):
        self.tmp.cleanup()

    def test_research_pack_runs(self):
        res = self.client.research_analysis_pack(
            file_path=self.path,
            sheet_name="Sheet0",
            profile="hlth3112",
        )
        self.assertTrue(res["success"])
        self.assertEqual(res["profile"], "hlth3112")
        self.assertIn("steps", res)
        self.assertIn("summary", res)
        self.assertGreater(res["summary"]["total_analyses"], 0)
        self.assertGreater(res["summary"]["succeeded"], 0)
        self.assertIn("qualitative", res["steps"])
        self.assertGreaterEqual(len(res["steps"]["qualitative"]), 1)

        # Confirm filters are applied in profile frequency/chi-square steps
        freq_step = res["steps"]["frequencies"][0]["result"]
        self.assertTrue(freq_step["success"])
        self.assertEqual(freq_step["allowed_values"], ["1", "2"])

    def test_research_pack_respects_require_formula_safe(self):
        # Safe workbook should pass strict gate
        ok = self.client.research_analysis_pack(
            file_path=self.path,
            sheet_name="Sheet0",
            profile="hlth3112",
            require_formula_safe=True,
        )
        self.assertTrue(ok["success"])
        self.assertTrue(ok["summary"]["require_formula_safe"])

        # Build unsafe workbook (unsupported IF). Bulk write neutralizes
        # formula-like text by design, so use the explicit formula entrypoint.
        unsafe = os.path.join(self.tmp.name, "unsafe.xlsx")
        self.client.write_spreadsheet(
            output_path=unsafe,
            headers=["A", "F", "M", "R", "Y", "AL", "AN", "AO", "AP", "AQ"],
            data=[
                [
                    1,
                    1,
                    "",
                    3,
                    4,
                    5,
                    "quote a",
                    "quote b",
                    "quote c",
                    "quote d",
                ],
                [2, 4, 5, 4, 5, 4, "quote e", "quote f", "quote g", "quote h"],
            ],
            sheet_name="Sheet0",
            overwrite_workbook=True,
            coerce_rows=True,
        )
        self.assertTrue(
            self.client.add_formula(unsafe, "C2", "=IF(1=1,4,2)", sheet_name="Sheet0")[
                "success"
            ]
        )
        blocked = self.client.research_analysis_pack(
            file_path=unsafe,
            sheet_name="Sheet0",
            profile="hlth3112",
            require_formula_safe=True,
        )
        self.assertFalse(blocked["success"])
        self.assertIn("Formula safety policy blocked execution", blocked["error"])

    def test_open_text_helpers(self):
        kw = self.client.open_text_keywords(
            file_path=self.path,
            column_letter="A",
            sheet_name="Sheet0",
            top_n=10,
            min_word_length=4,
        )
        self.assertTrue(kw["success"])
        self.assertGreaterEqual(kw["response_count"], 1)

        ex = self.client.open_text_extract(
            file_path=self.path,
            column_letter="A",
            sheet_name="Sheet0",
            limit=3,
            min_length=5,
        )
        self.assertTrue(ex["success"])
        self.assertLessEqual(ex["returned"], 3)


if __name__ == "__main__":
    unittest.main()
