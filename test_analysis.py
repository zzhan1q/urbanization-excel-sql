import csv
import unittest
from collections import defaultdict
from math import isclose

from build import FORMULA, ROOT, query, read_source
from openpyxl import load_workbook

MODE = "ends"
CODES = ("SP.URB.TOTL.IN.ZS",)
YEARS = (2010, 2023)


class AnalysisTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = read_source()
        cls.rows = {iso: (a, b, v) for _, iso, a, b, v in query(cls.data)}

    def test_sql_against_python(self):
        histories = defaultdict(dict)
        for iso, _, year, code, value in self.data:
            if value is not None:
                histories[iso][year, code] = value
        expected, pairs = {}, {}
        first, last = YEARS
        for iso, values in histories.items():
            a = values.get((first, CODES[0]))
            b = values.get((last, CODES[0]))
            if MODE == "threshold":
                series = sorted((year, v) for (year, _), v in values.items())
                if series[0][1] < 50:
                    for (year, x), (next_year, y) in zip(series, series[1:]):
                        if x < 50 <= y:
                            expected[iso] = (year, next_year, next_year - year)
                            break
                continue
            if MODE == "remit":
                series = [values.get((year, CODES[0])) for year in range(2019, 2024)]
                if None not in series and sum(series) >= 25:
                    expected[iso] = (
                        min(series),
                        max(series),
                        max(series) - min(series),
                    )
                continue
            if MODE in {"trade", "youth"}:
                a, b = b, values.get((last, CODES[1]))
            if a is None or b is None:
                continue
            if MODE == "recovery" and not any(
                values.get((yr, CODES[0]), a) < a for yr in (2020, 2021, 2022)
            ):
                continue
            if MODE in {"percent", "tourism"} and not a:
                continue
            if MODE == "rank":
                pairs[iso] = (a, b)
                continue
            metric = b - a
            if MODE == "trade":
                metric = b + a
            if MODE == "percent":
                metric = b / a * 100 - 100
            if MODE == "tourism":
                metric = b / a * 100
            expected[iso] = (a, b, metric)
        for iso, (a, b) in pairs.items():
            early = 1 + sum(x > a for x, _ in pairs.values())
            late = 1 + sum(y > b for _, y in pairs.values())
            expected[iso] = (early, late, early - late)
        self.assertTrue(expected)
        self.assertEqual(expected.keys(), self.rows.keys())
        for iso, wanted in expected.items():
            self.assertTrue(
                all(
                    isclose(x, y, abs_tol=1e-10) for x, y in zip(wanted, self.rows[iso])
                ),
                iso,
            )

    def test_workbook_formulas_filters_and_chart(self):
        with (ROOT / "summary.csv").open(encoding="utf-8-sig", newline="") as stream:
            csv_rows = {
                r["code"]: (
                    float(r["baseline"]),
                    float(r["comparison"]),
                    float(r["metric"]),
                )
                for r in csv.DictReader(stream)
            }
        self.assertEqual(csv_rows, self.rows)
        formulas = load_workbook(ROOT / "report.xlsx")
        cached = load_workbook(ROOT / "report.xlsx", data_only=True)
        self.addCleanup(formulas.close)
        self.addCleanup(cached.close)
        self.assertEqual(formulas.sheetnames, ["Результат", "Сверка", "Данные"])
        self.assertFalse(formulas.properties.creator)
        self.assertFalse(formulas.properties.lastModifiedBy)
        self.assertEqual(len(formulas["Результат"]._charts), 1)
        self.assertTrue(all(sheet.auto_filter.ref for sheet in formulas))
        for r in range(5, len(self.rows) + 5):
            report = cached["Результат"]
            self.assertEqual(
                formulas["Результат"][f"E{r}"].value, "=" + FORMULA.format(r=r)
            )
            a, b, metric = self.rows[report[f"B{r}"].value]
            self.assertTrue(isclose(report[f"E{r}"].value, metric, abs_tol=1e-10))
            self.assertTrue(isclose(report[f"F{r}"].value, metric, abs_tol=1e-10))
            self.assertTrue(isclose(report[f"C{r}"].value, a, abs_tol=1e-10))
            self.assertTrue(isclose(report[f"D{r}"].value, b, abs_tol=1e-10))
            self.assertEqual(report[f"G{r}"].value, 0)
        self.assertEqual(cached["Сверка"]["G1"].value, len(self.rows))
        self.assertEqual(cached["Сверка"]["G2"].value, 0)
        self.assertEqual(cached["Данные"].max_row, len(self.data) + 1)
