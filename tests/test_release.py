"""作者：Wangyuan；日期：2026-11-08。职责：复验培训文本身份、完整排名和日期/分箱边界。"""
import tempfile
import unittest
from pathlib import Path
import pandas as pd
from demo import synthetic_frame, example_config
from training import clean_ledger, analyze, parse_month_value


class TrainingRegressions(unittest.TestCase):
    def result(self, frame, config=None):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "synthetic.xlsx"
            frame.to_excel(path, index=False)
            cleaned, notes = clean_ledger(path)
            return analyze(cleaned, config or example_config(), notes)

    def test_leading_zero_identifiers_are_not_inferred_as_numbers(self) -> None:
        frame = synthetic_frame().iloc[:2].copy()
        frame["instructor"] = ["001", "1"]
        self.assertEqual(self.result(frame).kpis["instructors"], 2)

    def test_spreadsheet_rankings_keep_all_six_instructors(self) -> None:
        frame = pd.concat([synthetic_frame().iloc[:1]] * 6, ignore_index=True)
        frame["instructor"] = [f"Synthetic-{i}" for i in range(6)]
        result = self.result(frame)
        self.assertEqual(len(result.top_hours), 6)
        self.assertEqual(len(result.top_fees), 6)

    def test_close_distinct_bins_have_distinct_labels(self) -> None:
        config = example_config()
        config["quality_edges"] = [0, 1, 1.0000001, 1.0000002, 100]
        result = self.result(synthetic_frame(), config)
        self.assertEqual(result.quality_bins["label"].nunique(), 4)

    def test_complete_dates_accept_years_outside_twenty_first_century(self) -> None:
        for value in ["1999-12-31", "2100-01-01"]:
            self.assertFalse(pd.isna(parse_month_value(value)))
