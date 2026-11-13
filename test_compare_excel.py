"""
Тесты для compare_excel.py.

Запуск (из папки со скриптом):
    python -m pytest test_compare_excel.py -q

Зависимости: pip install pytest pandas openpyxl
"""

import sys
from datetime import datetime, time
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
import compare_excel as ce  # noqa: E402

OPT = ce.Options()


# ---------- вспомогательное ----------

def make(folder, name, data, sheet="Sheet1"):
    """Создаёт xlsx: data - словарь «столбец -> значения» или {лист: словарь}."""
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / name
    sheets = data if all(isinstance(v, dict) for v in data.values()) else {sheet: data}
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for title, columns in sheets.items():
            pd.DataFrame(columns).to_excel(writer, sheet_name=title, index=False)
    return path


def compare(tmp_path, data1, data2, check_balance=True, **options):
    f1 = make(tmp_path, "rep_datacore_1.xlsx", data1)
    f2 = make(tmp_path, "rep_noogadev_1.xlsx", data2)
    res = ce.compare_files(f1, f2, ce.Options(**options), ce.derive_names(f1, f2))
    if check_balance:
        assert statuses(res, "Баланс строк") == [ce.OK] * len(res.tables)
    return res


def statuses(res, name):
    return [c["Результат"] for c in res.checks if c["Проверка"] == name]


def comment(res, name):
    return " | ".join(c["Комментарий"] for c in res.checks if c["Проверка"] == name)


def run_main(monkeypatch, *args):
    monkeypatch.setattr(sys, "argv", ["compare_excel.py", *map(str, args)])
    return ce.main()


BASE = {"ID": ["a1", "b2", "c3", "d4"],
        "Имя": ["Иванов", "Петров", "Сидоров", "Кузнецов"],
        "Сумма": [10, 20, 30, 40]}


def changed(column, index, value, base=BASE):
    data = {k: list(v) for k, v in base.items()}
    data[column][index] = value
    return data


def reordered(order, base=BASE):
    return {k: [v[i] for i in order] for k, v in base.items()}


# ---------- значения ячеек ----------

@pytest.mark.parametrize("a, b, options, equal", [
    ("Иванов ", "Иванов", {}, True),
    ("   ", None, {}, True),
    ("", float("nan"), {}, True),
    (0.1 + 0.2, 0.3, {}, True),
    (1, 1.0, {}, True),
    (True, 1, {}, False),
    ("NA", None, {}, False),
    (10, 10.4, {}, False),
    (10, 10.4, {"tolerance": 0.5}, True),
    ("10", 10, {}, False),
    ("10", 10, {"text_numbers": True}, True),
    ("0810", 810, {"text_numbers": True}, True),
    ("1,5", 1.5, {"text_numbers": True}, True),
    ("12345678901234567", "12345678901234568", {"text_numbers": True}, False),
    ("ООО", "ооо", {}, False),
    ("ООО", "ооо", {"ignore_case": True}, True),
])
def test_values_equal(a, b, options, equal):
    assert ce.values_equal(a, b, ce.Options(**options)) is equal


@pytest.mark.parametrize("a, b, kind, reason", [
    (None, "x", "empty_a", ""),
    ("x", "  ", "empty_b", ""),
    ("0810", 810, "form", "потерян ведущий ноль"),
    ("810", 810, "form", "число вместо текста"),
    ("1,50", "1.5", "form", "разная запись числа"),
    ("ООО Ромашка", "ооо ромашка", "form", "отличается регистр"),
    ("ООО  Ромашка", "ООО Ромашка", "form", "отличаются пробелы"),
    ("ООО «Ромашка»", 'ООО "Ромашка"', "form", "отличаются кавычки"),
    (datetime(2026, 1, 2), "02.01.2026", "form", "дата вместо текста"),
    ("чай", "кофе", "value", "разные значения"),
    (10, 11, "value", "разные значения"),
])
def test_classify(a, b, kind, reason):
    assert ce.classify(a, b) == (kind, reason)


@pytest.mark.parametrize("value, expected", [
    (datetime(2026, 1, 1, 10, 5), True),
    (datetime(2026, 1, 1), False),
    (time(10, 5), True),
    ("01.01.2026 10:05:00", True),
    ("2026-01-01T10:05", True),
    ("10:05", True),
    ("01.01.2026", False),
    ("2024", False),
    ("RUB", False),
    (2024, False),
])
def test_is_stamp(value, expected):
    assert ce.is_stamp(value) is expected


# ---------- разбор параметров и мелкие функции ----------

def test_column_letters():
    assert [ce.column_index(x) for x in ("A", "Z", "AA", "ae")] == [0, 25, 26, 30]
    assert all(ce.column_index(ce.excel_letter(k)) == k for k in range(1000))


def test_parse_columns():
    assert ce.parse_columns(["AE", "AK:AM"]) == ((None, 30), (None, 36), (None, 37), (None, 38))
    assert ce.parse_columns(["Продажи!C:D"]) == (("Продажи", 2), ("Продажи", 3))
    assert ce.parse_columns(["A", "A:B"]) == ((None, 0), (None, 1))   # без повторов


@pytest.mark.parametrize("spec", ["1", "C:A", "A-B", "ABCD"])
def test_parse_columns_errors(spec):
    with pytest.raises(ValueError):
        ce.parse_columns([spec])


def test_column_sum():
    assert ce.column_sum([1, 2.5, "3", "1 000,5", None, "  ", "текст", True]) == (1007, 2)
    assert ce.column_sum([0.1, 0.2]) == (pytest.approx(0.3), 0)


def test_plural():
    forms = ("строка", "строки", "строк")
    assert [ce.plural(n, *forms) for n in (1, 2, 5, 11, 21, 112)] == [
        "строка", "строки", "строк", "строк", "строка", "строк"]


# ---------- имена файлов ----------

def test_derive_names():
    nm = ce.derive_names("pks07_datacore_20260101.xlsx", "pks07_noogadev_20260102.xlsx")
    assert (nm.a, nm.b, nm.prefix) == ("datacore", "noogadev", "pks07")
    nm = ce.derive_names("a/report.xlsx", "b/report.xlsx")
    assert (nm.a, nm.b, nm.prefix) == ("Файл 1", "Файл 2", "report")
    nm = ce.derive_names("x.xlsx", "y_z.xlsx", explicit=("старая", "новая"))
    assert (nm.a, nm.b) == ("старая", "новая")


def test_find_pairs(tmp_path):
    names = ["pks07_datacore_1.xlsx", "pks07_noogadev_1.xlsx",
             "pks08_datacore_1.xlsx", "pks08_noogadev_1.xlsx",
             "pks10_datacore_1.xlsx", "pks10_noogadev_1.xlsx",
             "pks09_datacore_1.xlsx",                       # без пары
             "pks07_1_differences.xlsx", "~$pks07_datacore_1.xlsx", ce.BATCH_SUMMARY]
    for n in names:
        (tmp_path / n).touch()
    pairs, unpaired, notes = ce.find_pairs(tmp_path, None)
    assert [(a.name, b.name) for a, b in pairs] == [
        tuple(names[0:2]), tuple(names[2:4]), tuple(names[4:6])]
    assert [f.name for f in unpaired] == ["pks09_datacore_1.xlsx"]
    assert notes == []


def test_find_pairs_two_reports_need_systems(tmp_path):
    """Ровно два отчёта: по именам не понять, что система, а что название отчёта."""
    for n in ("pks07_datacore_1.xlsx", "pks07_noogadev_1.xlsx",
              "pks08_datacore_1.xlsx", "pks08_noogadev_1.xlsx"):
        (tmp_path / n).touch()
    with pytest.raises(ValueError, match="--systems"):
        ce.find_pairs(tmp_path, None)
    pairs, unpaired, _ = ce.find_pairs(tmp_path, ("datacore", "noogadev"))
    assert len(pairs) == 2 and unpaired == []


def test_find_pairs_takes_latest_file(tmp_path):
    for n in ("pks07_datacore_1.xlsx", "pks07_datacore_2.xlsx", "pks07_noogadev_1.xlsx"):
        (tmp_path / n).touch()
    pairs, _, _ = ce.find_pairs(tmp_path, ("datacore", "noogadev"))
    assert len(pairs) == 1 and pairs[0][1].name == "pks07_noogadev_1.xlsx"


def test_find_pairs_not_found(tmp_path):
    (tmp_path / "one.xlsx").touch()
    with pytest.raises(ValueError):
        ce.find_pairs(tmp_path, None)


# ---------- сравнение: совпадающие файлы ----------

def test_identical(tmp_path):
    res = compare(tmp_path, BASE, BASE)
    assert not res.has_differences
    assert res.totals["exact"] == 4 and res.totals["changed"] == 0
    assert statuses(res, "Порядок строк") == [ce.OK]


def test_row_order_does_not_matter(tmp_path):
    res = compare(tmp_path, BASE, reordered([3, 1, 0, 2]))
    assert not res.has_differences
    assert statuses(res, "Порядок строк") == [ce.INFO]


def test_column_order_does_not_matter(tmp_path):
    res = compare(tmp_path, BASE, {k: BASE[k] for k in ("Сумма", "ID", "Имя")})
    assert not res.has_differences
    assert statuses(res, "Порядок столбцов") == [ce.INFO]


def test_spaces_and_float_noise_are_ignored(tmp_path):
    data2 = changed("Имя", 0, " Иванов  ")
    data2["Сумма"] = [10.000000000000002, 20, 30, 40]
    res = compare(tmp_path, BASE, data2)
    assert not res.has_differences
    assert statuses(res, "Пробелы в начале и конце текста") == [ce.WARN]


def test_single_sheets_with_different_names(tmp_path):
    res = compare(tmp_path, {"Лист1": BASE}, {"Data": BASE})
    assert not res.has_differences
    assert statuses(res, "Набор листов") == [ce.INFO]


# ---------- сравнение: расхождения ----------

def test_changed_cell_is_matched_by_id(tmp_path):
    res = compare(tmp_path, BASE, reordered([2, 0, 3, 1], changed("Сумма", 2, 31)))
    t = res.totals
    assert res.has_differences
    assert (t["exact"], t["changed"], t["only_a"], t["only_b"]) == (3, 1, 0, 0)
    assert (t["cells1"], t["cells2"], t["cells3"]) == (0, 0, 1)
    assert "по столбцу ID" in comment(res, "Способ сопоставления строк")
    d = res.details[0]
    assert (d["row_a"], d["row_b"], d["val_a"], d["val_b"]) == (4, 2, 30, 31)
    assert d["Столбец"] == "C: Сумма"


def test_empty_cell_groups(tmp_path):
    res = compare(tmp_path, BASE, changed("Имя", 1, None))
    assert (res.totals["cells1"], res.totals["cells2"], res.totals["cells3"]) == (1, 0, 0)
    res = compare(tmp_path, changed("Имя", 1, None), BASE)
    assert (res.totals["cells1"], res.totals["cells2"], res.totals["cells3"]) == (0, 1, 0)


def test_missing_and_extra_rows(tmp_path):
    short = {k: v[:3] for k, v in BASE.items()}
    res = compare(tmp_path, BASE, short)
    assert (res.totals["only_a"], res.totals["only_b"]) == (1, 0)
    assert statuses(res, "Количество строк с данными") == [ce.DIFF]
    res = compare(tmp_path, short, BASE)
    assert (res.totals["only_a"], res.totals["only_b"]) == (0, 1)


def test_duplicate_rows_are_counted(tmp_path):
    twice = {k: v + v[:1] for k, v in BASE.items()}
    res = compare(tmp_path, twice, BASE)
    assert res.has_differences
    assert (res.totals["exact"], res.totals["only_a"]) == (4, 1)
    assert statuses(res, "Повторяющиеся строки") == [ce.WARN]
    assert not compare(tmp_path, twice, reordered([4, 2, 0, 1, 3], twice)).has_differences


NO_ID = {"Город": ["Москва", "Москва", "Казань", "Казань"],
         "Товар": ["чай", "кофе", "чай", "кофе"],
         "Склад": ["А", "А", "Б", "Б"],
         "Кол-во": [5, 5, 7, 7]}


def test_matching_by_similarity_without_id(tmp_path):
    res = compare(tmp_path, NO_ID, reordered([3, 0, 2, 1], changed("Кол-во", 3, 9, NO_ID), ))
    t = res.totals
    assert (t["exact"], t["changed"], t["only_a"], t["only_b"]) == (3, 1, 0, 0)
    assert "по сходству" in comment(res, "Способ сопоставления строк")


def test_dissimilar_rows_are_not_paired(tmp_path):
    data2 = {k: list(v) for k, v in BASE.items()}
    for column, value in (("ID", "z9"), ("Имя", "Смирнов"), ("Сумма", 99)):
        data2[column][3] = value
    res = compare(tmp_path, BASE, data2)
    t = res.totals
    assert (t["exact"], t["changed"], t["only_a"], t["only_b"]) == (3, 0, 1, 1)


def test_type_difference_and_text_numbers(tmp_path):
    data2 = dict(BASE, Сумма=["10", "20", "30", "40"])
    res = compare(tmp_path, BASE, data2)
    assert res.totals["changed"] == 4
    assert "число вместо текста: 4" in comment(res, "Отличия в ячейках по причинам")
    assert statuses(res, "Типы данных в столбцах") == [ce.WARN]
    assert not compare(tmp_path, BASE, data2, text_numbers=True).has_differences


def test_na_text_is_not_empty(tmp_path):
    res = compare(tmp_path, changed("Имя", 0, "NA"), changed("Имя", 0, None))
    assert res.totals["cells1"] == 1


def test_tolerance(tmp_path):
    data2 = changed("Сумма", 0, 10.004)
    assert compare(tmp_path, BASE, data2).has_differences
    assert not compare(tmp_path, BASE, data2, tolerance=0.01).has_differences


# ---------- --key и --by-order ----------

def test_key(tmp_path):
    data2 = reordered([1, 0, 3, 2], changed("Имя", 0, "Иванова"))
    res = compare(tmp_path, BASE, data2, keys=["ID"])
    assert (res.totals["exact"], res.totals["changed"]) == (3, 1)
    assert res.details[0]["Ключ"] == "ID=a1"
    assert "По ключевым столбцам: ID" in comment(res, "Способ сопоставления строк")


def test_key_not_unique_and_missing(tmp_path):
    data = dict(BASE, ID=["a1", "a1", "c3", "d4"])
    res = compare(tmp_path, data, changed("Сумма", 1, 21, data), keys=["ID"])
    assert statuses(res, "Ключевые столбцы") == [ce.WARN]
    assert (res.totals["exact"], res.totals["changed"]) == (3, 1)
    res = compare(tmp_path, BASE, BASE, keys=["Нет такого"])
    assert statuses(res, "Ключевые столбцы") == [ce.WARN]
    assert not res.has_differences


def test_by_order(tmp_path):
    assert not compare(tmp_path, BASE, BASE, by_order=True).has_differences
    res = compare(tmp_path, BASE, reordered([1, 0, 2, 3]), by_order=True)
    assert (res.totals["exact"], res.totals["changed"]) == (2, 2)


# ---------- столбцы, отличающиеся во всех строках ----------

def with_column(name, value, base=BASE):
    return dict(base, **{name: [value] * len(base["ID"])})


def test_timestamp_column_is_technical(tmp_path):
    res = compare(tmp_path, with_column("Сформирован", datetime(2026, 1, 1, 10, 0)),
                  with_column("Сформирован", datetime(2026, 1, 1, 10, 5)))
    assert not res.has_differences
    assert [x[1] for x in res.tech] == ["D: Сформирован"] and res.whole == []
    assert statuses(res, "Технические столбцы") == [ce.TECH]
    assert res.totals["exact"] == 4            # остальные столбцы сопоставились
    assert "300 с" in ce.tech_comment(*res.tech[0][2:])


def test_text_timestamp_column_is_technical(tmp_path):
    res = compare(tmp_path, with_column("Сформирован", "01.01.2026 10:00:00"),
                  with_column("Сформирован", "01.01.2026 10:05:00"))
    assert not res.has_differences and len(res.tech) == 1


@pytest.mark.parametrize("v1, v2", [
    ("RUB", "USD"),
    ("2024", "2025"),
    (2024, 2025),
    (datetime(2026, 1, 1), datetime(2026, 1, 2)),      # дата без времени
    ("RUB", datetime(2026, 1, 1, 10, 0)),
])
def test_constant_column_is_a_difference(tmp_path, v1, v2):
    res = compare(tmp_path, with_column("Поле", v1), with_column("Поле", v2))
    assert res.has_differences
    assert [x[1] for x in res.whole] == ["D: Поле"] and res.tech == []
    assert statuses(res, "Столбцы, отличающиеся во всех строках") == [ce.DIFF]
    assert res.totals["exact"] == 4 and res.totals["const_real"] == 1


def test_constant_column_can_be_ignored(tmp_path):
    res = compare(tmp_path, with_column("Валюта", "RUB"), with_column("Валюта", "USD"),
                  ignore=("Валюта",))
    assert not res.has_differences
    assert statuses(res, "Исключённые столбцы") == [ce.INFO]


# ---------- структура ----------

def test_column_only_in_one_file(tmp_path):
    res = compare(tmp_path, BASE, dict(BASE, Доп=[1, 2, 3, 4]))
    assert statuses(res, "Набор столбцов") == [ce.DIFF]
    assert res.totals["structure"] == 1 and res.totals["exact"] == 4


def test_sheet_only_in_one_file(tmp_path):
    res = compare(tmp_path, {"Один": BASE, "Два": BASE}, {"Один": BASE, "Три": BASE})
    assert statuses(res, "Набор листов") == [ce.DIFF]
    assert res.totals["structure"] == 2 and len(res.tables) == 1


def test_selected_sheets(tmp_path):
    books = ({"Один": BASE, "Два": BASE}, {"Один": BASE, "Два": changed("Сумма", 0, 11)})
    assert compare(tmp_path, *books).has_differences
    assert not compare(tmp_path, *books, sheets=["Один"]).has_differences
    with pytest.raises(ValueError):
        compare(tmp_path, *books, sheets=["Нет такого"])


def test_skip_first_column(tmp_path):
    numbered = {"№": [1, 2, 3, 4], **BASE}
    data2 = {"№": [1, 2, 3, 4], **reordered([3, 2, 1, 0])}
    assert compare(tmp_path, numbered, data2).has_differences
    assert not compare(tmp_path, numbered, data2, skip_first_column=True).has_differences


# ---------- суммы ----------

def test_sums(tmp_path):
    o = {"sum_cols": ce.parse_columns(["C"])}
    res = compare(tmp_path, BASE, reordered([3, 2, 1, 0]), **o)
    assert statuses(res, "Суммы по указанным столбцам") == [ce.OK]
    entry = next(e for e in res.columns if "sum_a" in e)
    assert (entry["sum_a"], entry["sum_b"]) == (100, 100)

    res = compare(tmp_path, BASE, changed("Сумма", 0, 15), **o)
    assert statuses(res, "Суммы по указанным столбцам") == [ce.DIFF]
    entry = next(e for e in res.columns if "sum_a" in e)
    assert (entry["sum_a"], entry["sum_b"], entry["sum_differs"]) == (100, 105, True)


def test_sum_column_not_found(tmp_path):
    with pytest.raises(ValueError):
        compare(tmp_path, BASE, BASE, sum_cols=ce.parse_columns(["Z"]))
    with pytest.raises(ValueError):
        compare(tmp_path, BASE, BASE, sum_cols=ce.parse_columns(["Нет!A"]))


# ---------- запуск целиком: коды возврата и отчёт ----------

def pair(folder, data1=BASE, data2=BASE, prefix="rep"):
    return (make(folder, f"{prefix}_datacore_1.xlsx", data1),
            make(folder, f"{prefix}_noogadev_1.xlsx", data2))


def test_main_exit_codes_and_report_location(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    f1, f2 = pair(tmp_path / "same")
    assert run_main(monkeypatch, f1, f2) == 0
    report = tmp_path / "same" / "rep_1_differences.xlsx"     # рядом с первым файлом
    assert load_workbook(report).sheetnames == ["Итог"]
    assert "ДАННЫЕ СОВПАДАЮТ" in load_workbook(report)["Итог"]["A4"].value

    f1, f2 = pair(tmp_path / "diff", data2=changed("Сумма", 0, 11))
    assert run_main(monkeypatch, f1, f2) == 1
    sheet = load_workbook(tmp_path / "diff" / "rep_1_differences.xlsx")["Итог"]
    assert "РАСХОЖДЕНИЯ В ДАННЫХ" in sheet["A4"].value

    assert run_main(monkeypatch, f1, tmp_path / "нет_файла.xlsx") == 2
    assert list(tmp_path.glob("*.xlsx")) == []                # в текущую папку не пишем


def test_main_output_and_full(tmp_path, monkeypatch):
    f1, f2 = pair(tmp_path, data2=changed("Сумма", 0, 11))
    out = tmp_path / "out" / "my.xlsx"
    out.parent.mkdir()
    assert run_main(monkeypatch, f1, f2, "-o", out, "--full", "--names", "старая", "новая") == 1
    book = load_workbook(out)
    assert book.sheetnames == ["Итог", "Все проверки", "Описание различий",
                               "Различия по ячейкам", "Столбцы"]
    assert "старая ↔ новая" in book["Итог"]["A1"].value
    assert book["Различия по ячейкам"].max_row == 2


def test_main_constant_column_in_report(tmp_path, monkeypatch):
    f1, f2 = pair(tmp_path, with_column("Валюта", "RUB"), with_column("Валюта", "USD"))
    assert run_main(monkeypatch, f1, f2) == 1
    sheet = load_workbook(tmp_path / "rep_1_differences.xlsx")["Итог"]
    text = "\n".join(str(v) for row in sheet.iter_rows(values_only=True) for v in row if v)
    assert "столбцов отличается во всех строках: 1" in text
    assert "время формирования" not in text
    assert run_main(monkeypatch, f1, f2, "--ignore", "Валюта") == 0


@pytest.mark.parametrize("extra", [
    ["--similarity", "2"], ["--tolerance", "-1"], ["--header", "0"],
    ["--key", "ID", "--by-order"],
])
def test_main_rejects_bad_arguments(tmp_path, monkeypatch, extra):
    f1, f2 = pair(tmp_path)
    with pytest.raises(SystemExit) as error:
        run_main(monkeypatch, f1, f2, *extra)
    assert error.value.code == 2


def test_main_header_row(tmp_path, monkeypatch):
    shifted = {"Отчёт за январь": ["ID", *BASE["ID"]], "x": ["Имя", *BASE["Имя"]],
               "y": ["Сумма", *BASE["Сумма"]]}
    f1, f2 = pair(tmp_path, shifted, shifted)
    assert run_main(monkeypatch, f1, f2, "--header", "2", "--key", "ID") == 0


def test_main_dir(tmp_path, monkeypatch):
    pair(tmp_path, prefix="pks07")
    pair(tmp_path, data2=changed("Сумма", 0, 11), prefix="pks08")
    pair(tmp_path, prefix="pks10")
    make(tmp_path, "pks09_datacore_1.xlsx", BASE)
    assert run_main(monkeypatch, "--dir", tmp_path) == 1
    summary = pd.read_excel(tmp_path / ce.BATCH_SUMMARY)
    assert list(summary["Итог"]) == ["совпадают", "есть различия", "совпадают",
                                     "пара не найдена"]
    assert (tmp_path / "pks07_1_differences.xlsx").exists()
    # повторный запуск не принимает свои отчёты за исходные файлы
    assert run_main(monkeypatch, "--dir", tmp_path) == 1
    assert len(pd.read_excel(tmp_path / ce.BATCH_SUMMARY)) == 4


def test_main_dir_broken_pair_returns_error(tmp_path, monkeypatch):
    pair(tmp_path, prefix="pks07")
    (tmp_path / "pks08_datacore_1.xlsx").write_text("не excel")
    (tmp_path / "pks08_noogadev_1.xlsx").write_text("не excel")
    assert run_main(monkeypatch, "--dir", tmp_path, "--systems", "datacore", "noogadev") == 2
    summary = pd.read_excel(tmp_path / ce.BATCH_SUMMARY)
    assert summary["Итог"][0] == "совпадают" and summary["Итог"][1].startswith("ошибка")


def test_main_dir_rejects_sum(tmp_path, monkeypatch):
    pair(tmp_path)
    assert run_main(monkeypatch, "--dir", tmp_path, "--sum", "A") == 2
