import csv
import math
import sqlite3
from collections import Counter
from contextlib import closing
from pathlib import Path

from xlsxwriter import Workbook

TITLE = "Рост доли городского населения"
LIMIT = "Доля городского населения зависит от национального определения города. Рост доли не означает, что абсолютное сельское население уменьшилось."
INPUTS = ("2010, %", "2023, %")
OUTPUT = "Изменение, п. п."
FORMULA = "D{r}-C{r}"
ROOT = Path(__file__).resolve().parent


def read_source():
    with (ROOT / "source.csv").open(encoding="utf-8-sig", newline="") as file:
        raw = list(csv.DictReader(file))
    data, keys, names = [], set(), {}
    for r in raw:
        iso, name, year, code = r["iso3"], r["economy"], int(r["year"]), r["indicator"]
        value = float(r["value"]) if r["value"] else None
        key = (iso, year, code)
        if (
            key in keys
            or not 2010 <= year <= 2023
            or names.setdefault(iso, name) != name
        ):
            raise ValueError(f"Invalid observation: {key}")
        if value is not None and (not math.isfinite(value) or value < 0):
            raise ValueError(f"Invalid value: {key}")
        keys.add(key)
        data.append((iso, name, year, code, value))
    if not data:
        raise ValueError("Empty source")
    return data


def query(data):
    with closing(sqlite3.connect(":memory:")) as db:
        db.execute(
            "CREATE TABLE observations(iso TEXT, economy TEXT, year INTEGER, indicator TEXT, value REAL, PRIMARY KEY(iso, year, indicator))"
        )
        db.executemany("INSERT INTO observations VALUES (?, ?, ?, ?, ?)", data)
        return db.execute(
            (ROOT / "analysis.sql").read_text(encoding="utf-8")
        ).fetchall()


def workbook(data, result):
    coverage = Counter((r[3], r[2]) for r in data if r[4] is not None)
    expected = Counter((r[3], r[2]) for r in data)
    options = {"strings_to_formulas": False, "strings_to_urls": False}
    with Workbook(ROOT / "report.xlsx", options) as book:
        book.set_properties({"title": TITLE, "author": "", "comments": ""})
        banner = book.add_format(
            {"bold": True, "font_size": 17, "font_color": "#123541"}
        )
        header = book.add_format(
            {
                "bold": True,
                "bg_color": "#135566",
                "font_color": "#FFFFFF",
                "text_wrap": True,
                "valign": "vcenter",
            }
        )
        decimal = book.add_format({"num_format": "#,##0.00;[Red]-#,##0.00;0.00"})
        integer = book.add_format({"num_format": "#,##0"})
        note = book.add_format(
            {"text_wrap": True, "font_color": "#344E59", "valign": "vcenter"}
        )
        report = book.add_worksheet("Результат")
        report.hide_gridlines(2)
        report.merge_range("A1:F1", TITLE, banner)
        report.merge_range("A2:F2", LIMIT, note)
        report.set_row(1, 36)
        report.set_row(3, 32)
        report.set_column("A:A", 36)
        report.set_column("B:B", 10)
        report.set_column("C:F", 22, decimal)
        report.set_column("G:G", 17, decimal)
        report.write_row(
            3, 0, ["Экономика", "Код", *INPUTS, OUTPUT, "Из SQL", "Расхождение"], header
        )
        report.freeze_panes(4, 2)
        report.autofilter(3, 0, len(result) + 3, 6)
        for row, (name, iso, a, b, value) in enumerate(result, 4):
            report.write_row(row, 0, [name, iso, a, b])
            report.write_formula(
                row, 4, "=" + FORMULA.format(r=row + 1), decimal, value
            )
            report.write_number(row, 5, value)
            report.write_formula(row, 6, f"=ABS(E{row + 1}-F{row + 1})", decimal, 0)
        chart = book.add_chart({"type": "bar"})
        last = min(11, len(result) + 3)
        chart.add_series(
            {
                "categories": ["Результат", 4, 0, last, 0],
                "values": ["Результат", 4, 4, last, 4],
                "fill": {"color": "#1A7D80"},
                "border": {"none": True},
            }
        )
        chart.set_title({"name": OUTPUT + " · первые " + str(min(len(result), 8))})
        chart.set_y_axis({"reverse": True, "num_font": {"size": 9}})
        chart.set_legend({"none": True})
        report.insert_chart("I4", chart, {"x_scale": 1.15, "y_scale": 1.2})
        check = book.add_worksheet("Сверка")
        check.hide_gridlines(2)
        check.set_column("A:A", 32)
        check.set_column("B:D", 16, integer)
        check.write_row(0, 0, ["Показатель", "Год", "Значений", "Пустых"], header)
        for i, ((code, year), total) in enumerate(sorted(expected.items()), 1):
            count = coverage[(code, year)]
            check.write_row(i, 0, [code, year, count])
            check.write_formula(i, 3, f"={total}-C{i + 1}", integer, total - count)
        check.write("F1", "Строк в отчёте", header)
        check.write_formula(
            "G1", f"=COUNTA(Результат!B5:B{len(result) + 4})", integer, len(result)
        )
        check.write("F2", "Макс. расхождение", header)
        check.write_formula("G2", f"=MAX(Результат!G5:G{len(result) + 4})", decimal, 0)
        check.set_column("F:F", 23)
        check.set_column("G:G", 18)
        check.freeze_panes(1, 0)
        check.autofilter(0, 0, len(expected), 3)
        sheet = book.add_worksheet("Данные")
        sheet.set_column("A:A", 10)
        sheet.set_column("B:B", 36)
        sheet.set_column("C:C", 10)
        sheet.set_column("D:D", 26)
        sheet.set_column("E:E", 20, decimal)
        sheet.write_row(
            0, 0, ["Код", "Экономика", "Год", "Показатель", "Значение"], header
        )
        for i, r in enumerate(data, 1):
            sheet.write_row(i, 0, r)
        sheet.freeze_panes(1, 0)
        sheet.autofilter(0, 0, len(data), 4)


def main():
    data = read_source()
    result = query(data)
    if not result:
        raise ValueError("No comparable observations")
    with (ROOT / "summary.csv").open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["economy", "code", "baseline", "comparison", "metric"])
        writer.writerows(result)
    workbook(data, result)
    print(f"{len(result)} comparisons; {len(data)} source rows")


if __name__ == "__main__":
    main()
