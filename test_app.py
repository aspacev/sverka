"""
Тесты сайта сверок.

Запуск:  python -m pytest -q
"""

import io
import sqlite3
from datetime import date, datetime

import pandas as pd
import pytest
from openpyxl import load_workbook

import service
from app import create_app

DATES = [datetime(2026, 7, 5), datetime(2026, 7, 20), datetime(2026, 8, 10),
         datetime(2026, 8, 25), datetime(2026, 9, 3), datetime(2026, 9, 28)]
BASE = {"ID": ["a1", "b2", "c3", "d4", "e5", "f6"],
        "Клиент": ["Иванов", "Петров", "Сидоров", "Кузнецов", "Смирнов", "Попов"],
        "Дата": DATES,
        "Сумма": [10, 20, 30, 40, 50, 60]}


def xlsx(data):
    buffer = io.BytesIO()
    pd.DataFrame(data).to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer


def changed():
    """Июль: одна строка изменена. Сентябрь: одна изменена, одной нет."""
    data = {k: list(v)[:5] for k, v in BASE.items()}
    data["Сумма"][0] = 11
    data["Сумма"][4] = 55
    return data


@pytest.fixture
def site(tmp_path):
    app = create_app(tmp_path)
    app.config["TESTING"] = True
    client = app.test_client()
    client.post("/dict/systems/add", data={"names": "CRM\nБиллинг\nSAP"})
    client.post("/dict/reports/add", data={"system": 1, "names": "ПКС-07 Продажи\nПКС-08 Остатки"})
    client.post("/dict/reports/add", data={"system": 3, "names": "ПКС-40 Клиенты"})

    def query(sql, *params):
        conn = sqlite3.connect(tmp_path / "sverka.db")
        conn.row_factory = sqlite3.Row
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        return rows

    def compare(data_a=BASE, data_b=BASE, period=("2026-07-01", "2026-09-30"), report=1,
                system=1, **extra):
        form = {"report": report, "system": system, "from": period[0], "to": period[1],
                "file_a": (xlsx(data_a), "выгрузка datacore.xlsx"),
                "file_b": (xlsx(data_b), "выгрузка nooga.xlsx"), **extra}
        return client.post("/new", data=form, content_type="multipart/form-data")

    client.query, client.compare, client.data = query, compare, tmp_path
    return client


def page(client, url):
    response = client.get(url)
    assert response.status_code == 200, url
    return response.get_data(as_text=True)


# ---------- даты ----------

@pytest.mark.parametrize("value, expected", [
    (datetime(2026, 7, 5, 10, 30), date(2026, 7, 5)),
    (date(2026, 7, 5), date(2026, 7, 5)),
    ("05.07.2026", date(2026, 7, 5)),
    ("2026-07-05 10:30:00", date(2026, 7, 5)),
    ("2026-07", date(2026, 7, 1)),
    ("07.2026", date(2026, 7, 1)),
    ("Иванов", None), ("2026", None), (20260705, None), (None, None), ("", None),
])
def test_to_date(value, expected):
    assert service.to_date(value) == expected


def test_months():
    assert service.months_between(date(2025, 11, 15), date(2026, 2, 1)) == [
        "2025-11", "2025-12", "2026-01", "2026-02"]
    assert service.shift_month("2026-01", -1) == "2025-12"
    assert service.month_bounds("2024-02") == (date(2024, 2, 1), date(2024, 2, 29))


def test_find_date_column():
    period = (date(2026, 7, 1), date(2026, 9, 30))
    df = pd.DataFrame(dict(BASE, Сформирован=[datetime(2026, 9, 30, 12, 0)] * 6,
                           Договор=[datetime(2020, 1, k) for k in range(1, 7)]))
    assert service.find_date_column(df, *period) == "Дата"
    assert service.find_date_column(df.drop(columns="Дата"), *period) is None


# ---------- пустой сайт и справочники ----------

def test_empty_site(tmp_path):
    client = create_app(tmp_path).test_client()
    assert client.get("/").status_code == 302
    for url in ("/new", "/grid", "/history", "/dict"):
        assert client.get(url).status_code == 200
    assert "заполните справочники" in page(client, "/grid")
    assert client.get("/runs/1").status_code == 404


def test_dictionaries(site):
    html = page(site, "/dict")
    assert "Биллинг" in html and "ПКС-08 Остатки" in html
    site.post("/dict/systems/add", data={"names": "CRM\n1С"})          # повтор не добавится
    assert [r["name"] for r in site.query("SELECT name FROM systems ORDER BY id")] == [
        "CRM", "Биллинг", "SAP", "1С"]
    site.post("/dict/reports/2/save", data={"name": "ПКС-08 Склад", "date_column": "Дата остатка"})
    row = site.query("SELECT * FROM reports WHERE id = 2")[0]
    assert (row["name"], row["date_column"], row["date_source"]) == (
        "ПКС-08 Склад", "Дата остатка", "manual")
    site.post("/dict/reports/2/delete")
    assert len(site.query("SELECT * FROM reports")) == 2


def test_reports_belong_to_systems(site):
    # одно и то же название может быть у разных систем, но не дважды у одной
    site.post("/dict/reports/add", data={"system": 2, "names": "ПКС-07 Продажи"})
    site.post("/dict/reports/add", data={"system": 1, "names": "ПКС-07 Продажи"})
    site.post("/dict/reports/add", data={"names": "Без системы"})
    rows = site.query("SELECT system_id FROM reports WHERE name = 'ПКС-07 Продажи' ORDER BY id")
    assert [r[0] for r in rows] == [1, 2]
    assert len(site.query("SELECT * FROM reports")) == 4
    # в справочнике и в сетке видны только отчёты выбранной системы
    crm, sap = page(site, "/dict?system=1"), page(site, "/dict?system=3")
    assert 'value="ПКС-08 Остатки"' in crm and 'value="ПКС-40 Клиенты"' not in crm
    assert 'value="ПКС-40 Клиенты"' in sap and 'value="ПКС-08 Остатки"' not in sap
    grid = page(site, "/grid?system=3")
    assert 'data-name="пкс-40 клиенты"' in grid and 'data-name="пкс-08 остатки"' not in grid
    assert "<b>0 из 1</b>" in grid
    site.post("/dict/systems/add", data={"names": "Пустая"})
    assert "пока нет отчётов" in page(site, "/grid?system=4")
    # систему с отчётами удалить нельзя, пустую - можно
    site.post("/dict/systems/2/delete")
    site.post("/dict/systems/4/delete")
    assert [r[0] for r in site.query("SELECT id FROM systems ORDER BY id")] == [1, 2, 3]


def test_used_items_cannot_be_deleted(site):
    site.compare()
    site.post("/dict/systems/1/delete")
    site.post("/dict/reports/1/delete")
    assert len(site.query("SELECT * FROM systems")) == 3
    assert len(site.query("SELECT * FROM reports")) == 3


# ---------- сверка ----------

def test_run_with_month_breakdown(site):
    response = site.compare(BASE, changed())
    assert response.status_code == 302 and response.headers["Location"].endswith("/runs/1")
    run = site.query("SELECT * FROM runs")[0]
    assert (run["status"], run["by_month"], run["date_column"], run["unplaced"]) == (
        "diff", 1, "Дата", 0)
    assert (run["rows_a"], run["rows_b"], run["exact"], run["changed"], run["only_a"]) == (
        6, 5, 3, 2, 1)
    months = {m["month"]: (m["bad_rows"], m["partial"])
              for m in site.query("SELECT * FROM run_months")}
    assert months == {"2026-07": (1, 0), "2026-08": (0, 0), "2026-09": (2, 0)}
    diffs = site.query("SELECT kind, month FROM run_diffs ORDER BY id")
    assert [(d["kind"], d["month"]) for d in diffs] == [
        ("changed", "2026-07"), ("changed", "2026-09"), ("only_a", "2026-09")]
    assert site.query("SELECT date_column FROM reports WHERE id = 1")[0][0] == "Дата"

    html = page(site, "/runs/1")
    assert "Есть расхождения: 3 расхождения в строках" in html
    assert "Июль 2026:" in html and "Сентябрь 2026:" in html
    assert "выгрузка datacore.xlsx" in html
    july = page(site, "/runs/1?month=2026-07")
    assert july.count('class="first changed"') == 1
    assert '<td class="x">10</td>' in july and '<td class="x">11</td>' in july
    assert "расхождений в строках нет" in page(site, "/runs/1?month=2026-08")

    excel = site.get("/runs/1/excel")
    assert excel.status_code == 200 and excel.data[:2] == b"PK"
    book = load_workbook(io.BytesIO(excel.data))
    assert book.sheetnames == ["Итог", "Все проверки", "Описание различий",
                               "Различия по ячейкам", "Столбцы"]
    assert "ПКС-07 Продажи, CRM: datacore ↔ nooga" in book["Итог"]["A1"].value
    assert book["Различия по ячейкам"].max_row == 4          # заголовок и три расхождения
    assert (site.data / "uploads" / "1" / "a" / "выгрузка datacore.xlsx").is_file()


def test_empty_cell_is_marked(site):
    other = {k: list(v) for k, v in BASE.items()}
    other["Клиент"][1] = None
    site.compare(BASE, other)
    html = page(site, "/runs/1")
    assert '<td class="e">(пусто)</td>' in html and '<td class="x">Петров</td>' in html


def test_run_without_differences(site):
    site.compare()
    run = site.query("SELECT * FROM runs")[0]
    assert (run["status"], run["by_month"], run["exact"]) == ("ok", 1, 6)
    assert "Данные совпадают" in page(site, "/runs/1")
    assert page(site, "/grid?system=1&end=2026-09").count('<a class="cell ok"') >= 3


def test_run_without_date_column(site):
    no_date = {k: v for k, v in BASE.items() if k != "Дата"}
    other = {k: list(v) for k, v in no_date.items()}
    other["Сумма"][0] = 11
    site.compare(no_date, other)
    run = site.query("SELECT * FROM runs")[0]
    assert (run["status"], run["by_month"], run["unplaced"]) == ("diff", 0, 1)
    assert "не найден столбец с датой" in page(site, "/runs/1")
    html = page(site, "/grid?system=1&end=2026-09")
    assert 'colspan="3"' in html and "без разбивки по месяцам" in html


def test_manual_date_column(site):
    data = dict(BASE, Оплата=[datetime(2026, 8, 1)] * 3 + [datetime(2026, 9, 1)] * 3)
    other = {k: list(v) for k, v in data.items()}
    other["Сумма"][0] = 11                      # по «Дата» это июль, по «Оплата» - август
    site.post("/dict/reports/1/save", data={"name": "ПКС-07 Продажи", "date_column": "Оплата"})
    site.compare(data, other)
    months = {m["month"]: m["bad_rows"] for m in site.query("SELECT * FROM run_months")}
    assert months == {"2026-07": 0, "2026-08": 1, "2026-09": 0}
    assert site.query("SELECT date_column, date_source FROM reports WHERE id = 1")[0][1] == "manual"


def test_partial_month_and_rows_outside_period(site):
    site.compare(BASE, changed(), period=("2026-07-10", "2026-08-31"))
    run = site.query("SELECT * FROM runs")[0]
    months = {m["month"]: (m["bad_rows"], m["partial"])
              for m in site.query("SELECT * FROM run_months")}
    assert months == {"2026-07": (0, 1), "2026-08": (0, 0)}
    assert run["unplaced"] == 3                 # 5 июля и сентябрь - вне периода
    assert "Без даты или вне периода: 3" in page(site, "/runs/1")
    html = page(site, "/grid?system=1&end=2026-09")
    assert html.count('<a class="cell bad"') == 2 and ">!<" in html


def test_options_are_applied(site):
    other = dict(BASE, Сумма=[str(v) for v in BASE["Сумма"]])
    site.compare(BASE, other)
    site.compare(BASE, other, text_numbers="1")
    site.compare(BASE, changed(), ignore="Сумма", keys="ID")
    assert [r["status"] for r in site.query("SELECT status FROM runs ORDER BY id")] == [
        "diff", "ok", "diff"]
    assert site.query("SELECT changed, only_a FROM runs WHERE id = 3")[0][:] == (0, 1)


def test_broken_file_is_saved_as_error(site):
    form = {"report": 1, "system": 1, "from": "2026-07-01", "to": "2026-07-31",
            "file_a": (io.BytesIO(b"not excel"), "a.xlsx"), "file_b": (xlsx(BASE), "b.xlsx")}
    assert site.post("/new", data=form, content_type="multipart/form-data").status_code == 302
    assert site.query("SELECT status FROM runs")[0][0] == "error"
    assert "Сверка не выполнена" in page(site, "/runs/1")
    assert "ошибка" in page(site, "/history")
    assert '<a class="cell bad"' not in page(site, "/grid?system=1&end=2026-07")
    assert site.get("/runs/1/excel").status_code == 404


@pytest.mark.parametrize("change, message", [
    ({"system": ""}, "Выберите систему"),
    ({"report": 3}, "относится к другой системе"),
    ({"from": "2026-10-01"}, "позже даты"),
    ({"to": ""}, "Укажите даты периода"),
    ({"report": ""}, "Выберите отчёт"),
    ({"file_a": (io.BytesIO(b"x"), "a.csv")}, "должен быть Excel"),
])
def test_form_validation(site, change, message):
    form = {"report": 1, "system": 1, "from": "2026-07-01", "to": "2026-09-30",
            "file_a": (xlsx(BASE), "a.xlsx"), "file_b": (xlsx(BASE), "b.xlsx"), **change}
    response = site.post("/new", data=form, content_type="multipart/form-data")
    assert response.status_code == 200 and message in response.get_data(as_text=True)
    assert site.query("SELECT COUNT(*) FROM runs")[0][0] == 0


# ---------- сетка и история ----------

def test_grid(site):
    site.compare(BASE, changed())                                   # ПКС-07, июль - сентябрь
    html = page(site, "/grid?system=1&end=2026-10")
    row = html.split('data-name="пкс-07 продажи"')[1].split("</tr>")[0]
    assert row.count('<a class="cell bad"') == 2 and row.count('<a class="cell ok"') == 1
    assert row.count('<a class="cell no"') == 3                        # май, июнь, октябрь
    assert 'data-states="bad none ok"' in html
    assert "from=2026-10-01" in row and "to=2026-10-31" in row
    assert "<b>0 из 2</b>" in html
    # у другой системы клетки пустые; без параметра берётся система последней сверки
    assert '<a class="cell bad"' not in page(site, "/grid?system=3&end=2026-10")
    assert page(site, "/grid?end=2026-10").count('<a class="cell bad"') == 2
    assert "datacore ↔ nooga" in html
    assert page(site, "/grid?end=2026-10&n=12").count("<th>") == 13


def test_grid_shows_latest_run(site):
    site.compare(BASE, changed())
    site.compare(BASE, BASE, period=("2026-09-01", "2026-09-30"))   # сентябрь пересверен
    row = page(site, "/grid?system=1&end=2026-09").split('data-name="пкс-07 продажи"')[1]
    row = row.split("</tr>")[0]
    assert row.count('<a class="cell bad"') == 1 and "/runs/2" in row and "/runs/1" in row


def test_history_and_delete(site):
    site.compare(BASE, changed())
    site.compare(report=3, system=3)
    html = page(site, "/history")
    assert "Найдено: 2" in html and "3 расхождения" in html and "совпало" in html
    assert "Найдено: 1" in page(site, "/history?status=ok")
    assert "Найдено: 1" in page(site, "/history?report=1")
    assert "Найдено: 1" in page(site, "/history?system=3")
    assert site.post("/runs/1/delete").status_code == 302
    assert "Найдено: 1" in page(site, "/history")
    assert site.query("SELECT COUNT(*) FROM run_diffs")[0][0] == 0
    assert not (site.data / "uploads" / "1").exists()
    assert not (site.data / "reports" / "run_1.xlsx").exists()


def test_new_form_is_prefilled(site):
    html = page(site, "/new?report=2&system=1&from=2026-10-01&to=2026-10-31")
    assert '<option value="2" data-system="1" selected>' in html
    assert '<option value="1" selected>' in html and 'value="2026-10-01"' in html
