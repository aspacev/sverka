"""Запуск сверки и раскладка расхождений по месяцам."""

import calendar
import json
from datetime import date, datetime, time
from pathlib import Path

import compare_excel as ce

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август",
          "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
DATE_FORMATS = ((10, "%d.%m.%Y"), (10, "%Y-%m-%d"), (10, "%d/%m/%Y"), (7, "%Y-%m"), (7, "%m.%Y"))


# ---------- даты и месяцы ----------

def to_date(v):
    """Дата из ячейки (в том числе записанная текстом) или None."""
    if ce.is_empty(v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        s = v.strip()
        for length, fmt in DATE_FORMATS:
            if len(s) == length or (length == 10 and len(s) > 10 and s[10] in " T"):
                try:
                    return datetime.strptime(s[:length], fmt).date()
                except ValueError:
                    pass
    return None


def month_key(d):
    return f"{d.year:04d}-{d.month:02d}"


def month_bounds(key):
    year, month = int(key[:4]), int(key[5:7])
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def month_title(key):
    return MONTHS[int(key[5:7]) - 1]


def shift_month(key, delta):
    index = int(key[:4]) * 12 + int(key[5:7]) - 1 + delta
    return f"{index // 12:04d}-{index % 12 + 1:02d}"


def months_between(first, last):
    result, key, end = [], month_key(first), month_key(last)
    while key <= end:
        result.append(key)
        key = shift_month(key, 1)
    return result


# ---------- столбец с датой ----------

def find_date_column(df, period_from, period_to):
    """
    Ищет в листе столбец с датами: почти все значения - даты, значений больше
    одного и не меньше половины попадает в период сверки. Возвращает название.
    """
    best, best_score = None, 0.0
    for column in df.columns:
        values = [v for v in df[column].tolist()[:2000] if not ce.is_empty(v)]
        if len(values) < 1:
            continue
        dates = [d for d in map(to_date, values) if d is not None]
        if len(dates) < 0.8 * len(values):
            continue
        if len(values) > 1 and len({str(v) for v in values}) == 1:
            continue          # одно значение на весь файл: это не дата записи
        score = sum(period_from <= d <= period_to for d in dates) / len(dates)
        if score >= 0.5 and score > best_score:
            best, best_score = column, score
    return best


def resolve_column(df, name):
    """Столбец листа по названию из справочника (пробелы по краям не важны)."""
    wanted = str(name).strip().casefold()
    for column in df.columns:
        if str(column).strip().casefold() == wanted:
            return column
    return None


# ---------- значения для показа на сайте ----------

def show(v):
    if ce.is_empty(v):
        return None
    if isinstance(v, datetime):
        plain = (v.hour, v.minute, v.second) == (0, 0, 0)
        return v.strftime("%d.%m.%Y" if plain else "%d.%m.%Y %H:%M:%S")
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, bool):
        return "ИСТИНА" if v else "ЛОЖЬ"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:.12g}"
    return str(v)


def dump(value):
    return json.dumps(value, ensure_ascii=False, default=show)


# ---------- сверка ----------

def execute(conn, data_dir, run_id, file_a, file_b, system_a, system_b, report,
            period_from, period_to, options, title=None):
    """
    Сверяет два файла и записывает результат в базу. Возвращает статус.
    system_a и system_b - названия сторон сверки, title - заголовок Excel-отчёта.
    """
    try:
        result = _compare(conn, data_dir, run_id, file_a, file_b, system_a, system_b,
                          report, period_from, period_to, options, title)
    except PermissionError as e:
        result = {"status": "error", "error_text": f"Нет доступа к файлу {e.filename}"}
    except Exception as e:
        text = str(e) or type(e).__name__
        result = {"status": "error", "error_text": text[:2000]}
    columns = ", ".join(f"{k} = ?" for k in result)
    conn.execute(f"UPDATE runs SET {columns} WHERE id = ?", (*result.values(), run_id))
    conn.commit()
    return result["status"]


def _compare(conn, data_dir, run_id, file_a, file_b, system_a, system_b, report,
             period_from, period_to, options, title=None):
    o = ce.Options(**options)
    nm = ce.Names(system_a, system_b, system_a, system_b, title or report["name"])
    res = ce.compare_files(file_a, file_b, o, nm)
    book1, book2 = res.books
    first_row = o.header + 2

    reports_dir = Path(data_dir) / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)
    ce.write_report(reports_dir / f"run_{run_id}.xlsx", res, file_a, file_b, nm, o)

    # --- столбец с датой на каждом листе ---
    sheets, found = {}, None

    def sheet_info(label):
        if label not in sheets:
            if label in book1 or label in book2:
                frames = (book1.get(label), book2.get(label))
            elif len(book1) == 1 and len(book2) == 1:      # один лист с разными названиями
                frames = (next(iter(book1.values())), next(iter(book2.values())))
            else:
                frames = (None, None)
            info = []
            for df in frames:
                column = None
                if df is not None:
                    if report["date_source"] == "manual" and report["date_column"]:
                        column = resolve_column(df, report["date_column"])
                    else:
                        column = find_date_column(df, period_from, period_to)
                info.append((df, column))
            sheets[label] = info
        return sheets[label]

    for tb in res.tables:
        for _, column in sheet_info(tb["sheet"]):
            if column is not None:
                found = found or str(column).strip()

    def month_of(label, row_a, row_b):
        for (df, column), row in zip(sheet_info(label), (row_a, row_b)):
            if df is None or column is None or not isinstance(row, int):
                continue
            index = row - first_row
            if 0 <= index < len(df):
                d = to_date(df[column].iloc[index])
                if d is not None and period_from <= d <= period_to:
                    return month_key(d)
            return None
        return None

    # --- строки с расхождениями по месяцам ---
    per_month, unplaced, seen = {}, 0, set()
    for d in res.details:
        row_a, row_b = d["row_a"], d["row_b"]
        if not isinstance(row_a, int) and not isinstance(row_b, int):
            continue                                       # отличие в структуре
        key = (d["Лист"], row_a, row_b)
        if key in seen:
            continue
        seen.add(key)
        month = month_of(*key)
        if month is None:
            unplaced += 1
        else:
            per_month[month] = per_month.get(month, 0) + 1

    conn.execute("DELETE FROM run_months WHERE run_id = ?", (run_id,))
    conn.execute("DELETE FROM run_diffs WHERE run_id = ?", (run_id,))
    for month in months_between(period_from, period_to):
        first, last = month_bounds(month)
        partial = period_from > first or period_to < last
        conn.execute("INSERT INTO run_months VALUES (?, ?, ?, ?)",
                     (run_id, month, per_month.get(month, 0), int(partial)))

    # --- сами расхождения для показа на сайте ---
    tables = []
    for tb in res.tables:
        label = tb["sheet"]
        tables.append({"sheet": label, "columns": tb["columns"], "tech": tb["tech"],
                       "whole": tb["whole"], "changed": tb["changed"],
                       "only_n": tb["only_n"]})
        for row_a, row_b, vals_a, vals_b, kinds, hint in tb["pairs"]:
            conn.execute(
                "INSERT INTO run_diffs (run_id, sheet, kind, row_a, row_b, month, vals_a, "
                "vals_b, kinds, hint) VALUES (?, ?, 'changed', ?, ?, ?, ?, ?, ?, ?)",
                (run_id, label, row_a, row_b, month_of(label, row_a, row_b),
                 dump([show(v) for v in vals_a]), dump([show(v) for v in vals_b]),
                 dump(kinds), hint))
        for system, row, values, hint in tb["only"]:
            in_a = system == nm.a
            row_a, row_b = (row, None) if in_a else (None, row)
            conn.execute(
                "INSERT INTO run_diffs (run_id, sheet, kind, row_a, row_b, month, vals_a, "
                "vals_b, kinds, hint) VALUES (?, ?, ?, ?, ?, ?, ?, ?, '{}', ?)",
                (run_id, label, "only_a" if in_a else "only_b", row_a, row_b,
                 month_of(label, row_a, row_b),
                 dump([show(v) for v in values]) if in_a else None,
                 None if in_a else dump([show(v) for v in values]), hint))

    if report["date_source"] != "manual" and found and found != report["date_column"]:
        conn.execute("UPDATE reports SET date_column = ? WHERE id = ?", (found, report["id"]))

    t = res.totals
    summary = {
        "tables": tables,
        "checks": [c for c in res.checks if c["Результат"] != ce.OK or c["Комментарий"]],
        "whole": [(s, c, show(v1), show(v2)) for s, c, v1, v2 in res.whole],
        "tech": [(s, c, show(v1), show(v2)) for s, c, v1, v2 in res.tech],
        "structure": [d for d in res.descriptions if d["Категория"] == "Структура"],
    }
    return {
        "status": "diff" if res.has_differences else "ok", "error_text": None,
        "rows_a": t["rows_a"], "rows_b": t["rows_b"], "exact": t["exact"],
        "changed": t["changed"], "only_a": t["only_a"], "only_b": t["only_b"],
        "cells": t["cells1"] + t["cells2"] + t["cells3"],
        "structure": t["structure"] + t["const_real"],
        "by_month": int(found is not None), "date_column": found,
        "unplaced": unplaced, "summary": dump(summary),
    }
