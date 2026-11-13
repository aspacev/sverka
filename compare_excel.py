#!/usr/bin/env python3
"""
Сверка данных в двух Excel-файлах с одинаковой структурой.

Структура (листы, столбцы) определяется автоматически, поэтому скрипт
подходит для любой пары файлов. Порядок строк, столбцов и листов
в файлах может быть разным. Сравниваются все столбцы, включая столбец A.

Запуск:
    python compare_excel.py файл1.xlsx файл2.xlsx
    python compare_excel.py --dir ПАПКА            (все пары файлов в папке)

Примеры:
    python compare_excel.py pks07_datacore_1.xlsx pks07_noogadev_1.xlsx
    python compare_excel.py a.xlsx b.xlsx -o report.xlsx
    python compare_excel.py a.xlsx b.xlsx --key ID
    python compare_excel.py a.xlsx b.xlsx --skip-first-column --text-numbers
    python compare_excel.py a.xlsx b.xlsx --sum AH:AQ          (суммы по столбцам)
    python compare_excel.py a.xlsx b.xlsx --sum AE AH AK:AM
    python compare_excel.py --dir C:\\Reports --systems datacore noogadev

Первый файл считается эталоном: отчёт показывает, что с его данными произошло
во втором файле. Каждое расхождение относится к одной из трёх групп:
    Есть в файле 1, нет в файле 2    данные потеряны (пустая ячейка или нет строки)
    Нет в файле 1, есть в файле 2    данные лишние
    Значения разные                  заполнено в обоих файлах, но по-разному

Отчёт (по умолчанию <общая часть имени>_differences.xlsx рядом с первым
файлом) - один лист «Итог»:
    - вывод: совпадают данные или нет, сколько строк и ячеек расходится;
    - таблица расхождений: для каждой отличающейся записи две строки (из первого
      и из второго файла) со всеми столбцами, отличающиеся ячейки подсвечены;
    - строки, которые есть только в одном из файлов;
    - технические отличия (столбец с одной отметкой времени на весь файл,
      например время формирования отчёта): показаны отдельно и на итог не влияют;
    - замечания, проверки без замечаний и справочная информация.
С параметром --full добавляются подробные листы: «Все проверки», «Описание
различий», «Различия по ячейкам», «Столбцы». Они же добавляются сами, если
расхождений больше, чем помещается в таблицу на листе «Итог».

Как сопоставляются строки (порядок строк не важен):
    1. Полностью одинаковые строки считаются совпавшими (с учётом повторов).
    2. Оставшиеся строки связываются по столбцу-идентификатору, если в данных
       есть столбец без повторов (скрипт находит его сам).
    3. Остальные связываются по сходству: совпадает не меньше половины ячеек.
    4. Строки без пары попадают в отчёт как «только в одном из файлов».
    --key СТОЛБЕЦ   сопоставлять строки только по указанным столбцам
    --by-order      сравнивать строго по порядку (2-я строка со 2-й и т.д.)

Суммирование: скрипт сам ничего не суммирует. Столбцы для подсчёта сумм
задаются параметром --sum буквами Excel (как в первом файле): отдельные буквы
и диапазоны через двоеточие. Суммы обоих файлов показываются на листе
«Столбцы»; расхождение сумм считается различием. Если в файле несколько
листов, буква применяется к тем листам, где такой столбец есть; чтобы указать
конкретный лист, пишите «Лист!Столбец», например: --sum "Продажи!C:E".

Что при сравнении считается одинаковым:
    - пробелы в начале и конце текста не учитываются («Иванов » = «Иванов»);
      сколько таких ячеек в каждом файле, показано в «Сводке»;
    - пустая ячейка и ячейка из одних пробелов - одно и то же;
    - числа сравниваются с точностью до 15 значащих цифр, чтобы не ловить шум
      вычислений (0.30000000000000004 = 0.3); больший допуск задаётся --tolerance;
    - текст «NA», «NULL», «None» и подобный - это текст, а не пустая ячейка.

Если столбец в каждом файле заполнен одним значением, а между файлами эти
значения разные, он отмечается одной записью и не мешает сопоставлению
остальных данных. Техническим отличием (на итог не влияет) считается только
отметка времени - дата со временем или время, например время формирования
отчёта. Любое другое значение (валюта, год, дата без времени) - расхождение;
служебный столбец можно исключить параметром --ignore.

Код возврата: 0 - различий нет, 1 - различия есть, 2 - ошибка (в режиме --dir -
если не удалось сравнить хотя бы одну пару).

Зависимости: pip install pandas openpyxl
Для быстрого чтения больших файлов (в несколько раз быстрее):
             pip install python-calamine
Если python-calamine не установлена, файлы читаются обычным способом.
"""

import argparse
import math
import re
import sys
from collections import Counter, defaultdict, deque
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path

import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.formatting.rule import FormulaRule
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OK, DIFF, WARN, INFO = "совпадает", "есть различия", "обратить внимание", "справочно"
TECH = "техническое отличие"   # на итог не влияет
MAX_TABLE_PAIRS = 2000         # сколько записей показывать в таблице на листе «Итог»
# Группа расхождения по виду отличия в ячейке: 1 - есть в файле 1, нет в файле 2;
# 2 - нет в файле 1, есть в файле 2; 3 - значения разные
GROUP_OF = {"empty_b": 1, "empty_a": 2, "value": 3, "form": 3}
MAX_REPORT_ROWS = 300_000      # ограничение на число строк в листах отчёта
TOO_COMMON = 100               # значения, встречающиеся чаще, не помогают искать пару
NUMBER_RE = re.compile(r"[+-]?\d+([.,]\d+)?")
QUOTES_RE = re.compile("[\"'«»“”„‘’`]")
BATCH_SUMMARY = "summary_all_reports.xlsx"
# Дата со временем или время: так выглядит отметка о формировании отчёта
STAMP_RE = re.compile(
    r"(\d{4}-\d{2}-\d{2}|\d{2}[./]\d{2}[./]\d{4})[ T]\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?"
    r"(Z|[+-]\d{2}:?\d{2})?|\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?")


@dataclass
class Options:
    keys: list = None
    by_order: bool = False
    header: int = 0              # номер строки заголовков, считая с нуля
    tolerance: float = 0.0
    ignore_case: bool = False
    text_numbers: bool = False
    similarity: float = 0.5
    ignore: tuple = ()
    skip_first_column: bool = False
    sheets: list = None
    sum_cols: tuple = ()         # номера столбцов (с нуля), по которым считать суммы
    full: bool = False           # добавлять в отчёт подробные листы


@dataclass
class Names:
    """Как называть файлы в отчёте: a/b - в заголовках, a_in/b_in - после «в»."""
    a: str = "Файл 1"
    b: str = "Файл 2"
    a_in: str = "файле 1"
    b_in: str = "файле 2"
    prefix: str = ""             # общая часть имён файлов (название отчёта)


# ---------- значения ячеек ----------

def is_empty(v):
    """Пустая ячейка: None, NaN или строка из пробелов."""
    if v is None or v is pd.NaT or v is pd.NA:
        return True
    if isinstance(v, str):
        return not v.strip()
    if isinstance(v, float):
        return v != v
    return False


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_stamp(v):
    """Отметка времени: дата со временем или время (в ячейке или текстом)."""
    if isinstance(v, datetime):
        return bool(v.hour or v.minute or v.second or v.microsecond)
    if isinstance(v, time):
        return True
    return isinstance(v, str) and bool(STAMP_RE.fullmatch(v.strip()))


def canon(v, o):
    """Приводит значение ячейки к виду, удобному для сравнения."""
    if is_empty(v):
        return None
    if isinstance(v, bool):
        return ("bool", v)
    if isinstance(v, str):
        v = v.strip()
        # Текст, похожий на число, по желанию сравниваем как число. Длинные
        # номера (больше 15 цифр) оставляем текстом, чтобы не терять точность.
        if (o.text_numbers and NUMBER_RE.fullmatch(v)
                and sum(ch.isdigit() for ch in v) <= 15):
            return clean_float(float(v.replace(",", ".")))
        return v.casefold() if o.ignore_case else v
    if is_number(v):
        try:
            return clean_float(float(v))
        except OverflowError:
            return v
    return v


def clean_float(f):
    """
    Оставляет 15 значащих цифр: убирает шум двоичной арифметики
    (0.30000000000000004 -> 0.3), не трогая значимые разряды.
    """
    return f if f.is_integer() else float(f"{f:.15g}")


def values_equal(a, b, o):
    a, b = canon(a, o), canon(b, o)
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) <= o.tolerance   # шум вычислений уже убран в canon
    return a == b


def type_label(v):
    if isinstance(v, bool):
        return "логическое"
    if isinstance(v, str):
        return "текст"
    if is_number(v):
        return "число"
    if isinstance(v, (datetime, date)):
        return "дата"
    if isinstance(v, time):
        return "время"
    return type(v).__name__


def as_number(text):
    """Число из текста или None (длинные номера числом не считаем)."""
    text = text.strip()
    if NUMBER_RE.fullmatch(text) and sum(ch.isdigit() for ch in text) <= 15:
        return float(text.replace(",", "."))
    return None


def has_leading_zero(text):
    digits = text.strip().lstrip("+-")
    return len(digits) > 1 and digits[0] == "0" and digits[1] not in ".,"


def classify(a, b):
    """
    Чем отличаются два неравных значения. Возвращает (вид, причина):
    вид - empty_a / empty_b / form (та же суть, другая запись) / value.
    """
    if is_empty(a):
        return "empty_a", ""
    if is_empty(b):
        return "empty_b", ""

    if is_number(a) != is_number(b) and (isinstance(a, str) or isinstance(b, str)):
        text, num = (a, b) if isinstance(a, str) else (b, a)
        if is_number(num) and as_number(text) == float(num):
            if has_leading_zero(text):
                return "form", "потерян ведущий ноль"
            return "form", "число вместо текста"

    if isinstance(a, str) and isinstance(b, str):
        s1, s2 = a.strip(), b.strip()
        n1, n2 = as_number(s1), as_number(s2)
        if n1 is not None and n1 == n2:
            if has_leading_zero(s1) != has_leading_zero(s2):
                return "form", "потерян ведущий ноль"
            return "form", "разная запись числа"
        if s1.casefold() == s2.casefold():
            return "form", "отличается регистр"
        if " ".join(s1.split()) == " ".join(s2.split()):
            return "form", "отличаются пробелы"
        if QUOTES_RE.sub("", s1) == QUOTES_RE.sub("", s2):
            return "form", "отличаются кавычки"

        def loose(s):
            return " ".join(QUOTES_RE.sub("", s).split()).casefold()
        if loose(s1) == loose(s2):
            return "form", "отличаются кавычки, пробелы или регистр"

    for d, s in ((a, b), (b, a)):
        if isinstance(d, (datetime, date)) and isinstance(s, str):
            forms = {str(d), d.isoformat(), d.strftime("%Y-%m-%d"),
                     d.strftime("%d.%m.%Y"), d.strftime("%Y-%m-%d %H:%M:%S"),
                     d.strftime("%d.%m.%Y %H:%M:%S")}
            if s.strip() in forms:
                return "form", "дата вместо текста"

    return "value", "разные значения"


# ---------- вспомогательное для текста ----------

def plural(n, one, few, many):
    n = abs(n) % 100
    if 11 <= n <= 14:
        return many
    return {1: one}.get(n % 10, few if 2 <= n % 10 <= 4 else many)


def column_index(letters):
    """Буква столбца Excel -> номер с нуля: A -> 0, Z -> 25, AA -> 26."""
    index = 0
    for ch in letters.upper():
        index = index * 26 + ord(ch) - 64
    return index - 1


def parse_columns(specs):
    """
    Разбирает «AE AH AK:AM» и «Лист!C:E» в список пар (лист или None, номер
    столбца с нуля).
    """
    result = []
    for arg in specs:
        sheet, _, cols = arg.rpartition("!")
        sheet = sheet.strip().strip("'") or None
        for spec in cols.replace(",", " ").split():
            m = re.fullmatch(r"([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?", spec)
            if not m:
                raise ValueError(
                    f"Не понял столбец «{spec}». Укажите буквы столбцов Excel, например: "
                    "--sum AE AH, диапазон --sum AK:AM или с листом --sum \"Лист1!C:E\"")
            first = column_index(m[1])
            last = column_index(m[2]) if m[2] else first
            if last < first:
                raise ValueError(f"В диапазоне «{spec}» первый столбец стоит правее второго")
            result += [(sheet, k) for k in range(first, last + 1)
                       if (sheet, k) not in result]
    return tuple(result)


def to_number(v):
    """Число из ячейки (в том числе записанное текстом) или None."""
    if is_number(v):
        return float(v)
    if isinstance(v, str):
        t = v.strip().replace("\u00a0", "").replace(" ", "")
        if NUMBER_RE.fullmatch(t):
            return float(t.replace(",", "."))
    return None


def column_sum(values):
    """Сумма столбца и число непустых ячеек, которые не являются числами."""
    numbers, bad = [], 0
    for v in values:
        if is_empty(v):
            continue
        n = to_number(v)
        if n is None:
            bad += 1
        else:
            numbers.append(n)
    total = math.fsum(numbers)
    return (int(total) if total == int(total) else total), bad


def number_text(x):
    """Число для показа в тексте: 1 234 567.89"""
    if x is None:
        return "нет"
    if isinstance(x, int):
        return f"{x:,}".replace(",", " ")
    return f"{x:,.6f}".rstrip("0").rstrip(".").replace(",", " ")


def excel_letter(index):
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def display(col):
    """Название столбца для отчёта."""
    m = re.fullmatch(r"Unnamed: (\d+)", str(col))
    if m:
        return f"столбец {excel_letter(int(m[1]))} (без заголовка)"
    return str(col)


def listing(items, limit=8):
    items = [str(x) for x in items]
    if len(items) <= limit:
        return ", ".join(items)
    return ", ".join(items[:limit]) + f" и ещё {len(items) - limit}"


def brief(v, limit=40):
    s = str(v)
    return s if len(s) <= limit else s[:limit - 3] + "..."


def describe_row(columns, row, limit=250):
    text = "; ".join(f"{display(c)}={v}" for c, v in zip(columns, row)
                     if not is_empty(v))
    return text if len(text) <= limit else text[:limit - 3] + "..."


def cap(text):
    return text[:1].upper() + text[1:]


# ---------- результат сравнения ----------

@dataclass
class Result:
    checks: list = field(default_factory=list)        # лист «Сводка»
    descriptions: list = field(default_factory=list)  # «Описание различий»
    details: list = field(default_factory=list)       # «Различия по ячейкам»
    columns: list = field(default_factory=list)       # «Столбцы»
    totals: Counter = field(default_factory=Counter)
    sum_used: set = field(default_factory=set)        # какие столбцы --sum нашлись
    tables: list = field(default_factory=list)        # таблицы расхождений для «Итога»
    tech: list = field(default_factory=list)          # технические отличия
    whole: list = field(default_factory=list)         # столбцы, разные во всех строках
    books: tuple = ()                                 # прочитанные файлы (для сайта)

    @property
    def has_differences(self):
        return any(c["Результат"] == DIFF for c in self.checks)

    def check(self, sheet, name, status, a="", b="", comment=""):
        self.checks.append({"Лист": sheet, "Проверка": name, "Результат": status,
                            "a": a, "b": b, "Комментарий": comment})

    def describe(self, sheet, category, where, what):
        self.descriptions.append({"Лист": sheet, "Категория": category,
                                  "Где": where, "Что отличается": what})

    def detail(self, sheet, group, key="", row_a=None, row_b=None, column="",
               val_a=None, val_b=None, reason=""):
        """group: 1 - есть в файле 1, нет в файле 2; 2 - наоборот; 3 - значения разные."""
        self.details.append({"Лист": sheet, "group": group, "Ключ": key,
                             "row_a": row_a, "row_b": row_b, "Столбец": column,
                             "val_a": val_a, "val_b": val_b, "Причина": reason})


def column_list(columns):
    return "\n".join(f"  {excel_letter(i)}: {display(c)}" for i, c in enumerate(columns))


# ---------- сравнение одного листа ----------

def looks_like_numbering(values):
    """Числа подряд, начиная с 0 или 1 (значения уже приведены canon)."""
    values = [v for v in values if v is not None]
    if len(values) < 2 or not all(isinstance(v, float) and v == int(v) for v in values):
        return False
    return values[0] in (0, 1) and all(y - x == 1 for x, y in zip(values, values[1:]))


def same_cells(n1, n2, skip=-1):
    """Сколько заполненных ячеек совпадает в двух строках (кроме столбца skip)."""
    return sum(1 for c, (a, b) in enumerate(zip(n1, n2))
               if a is not None and a == b and c != skip)


def find_id_column(columns, norm1, norm2, active1, active2, left1, left2):
    """
    Ищет столбец-идентификатор: без повторов и почти без пустых в обоих файлах,
    значения - текст или целые числа. Если подходящих столбцов несколько (или
    единственный - на самом деле сумма), выбирается тот, по которому связанные
    строки больше всего похожи друг на друга. Возвращает индекс или None.
    """
    def suitable(norm, active, c):
        values = [norm[i][c] for i in active]
        filled = [v for v in values if v is not None]
        if len(filled) < 2 or len(filled) < 0.9 * len(values):
            return False
        if not all(isinstance(v, str) or (isinstance(v, float) and v.is_integer())
                   for v in filled):
            return False
        return len(set(filled)) == len(filled) and not looks_like_numbering(values)

    found = [c for c, name in enumerate(columns)
             if not str(name).startswith("Unnamed: ")
             and suitable(norm1, active1, c) and suitable(norm2, active2, c)]
    if len(columns) == 1:
        return found[0] if found else None

    best, best_score = None, 0
    sample = left1[:2000]
    for c in found:
        by_id = {norm2[j][c]: j for j in left2 if norm2[j][c] is not None}
        score = 0
        for i in sample:
            j = by_id.get(norm1[i][c])
            if j is not None:
                score += same_cells(norm1[i], norm2[j], skip=c)
        if score > best_score:   # при равенстве остаётся столбец левее
            best, best_score = c, score
    return best


def pair_by_similarity(left1, left2, norm1, norm2, rows1, rows2, o):
    """Связывает оставшиеся строки, у которых совпадает достаточная доля ячеек."""
    pairs, used1, used2 = [], set(), set()

    def ratio(i, j):
        same = total = 0
        for a, b in zip(rows1[i], rows2[j]):
            if is_empty(a) and is_empty(b):
                continue
            total += 1
            same += values_equal(a, b, o)
        return same / total if total else 0.0

    def take(i, j):
        used1.add(i)
        used2.add(j)
        pairs.append((i, j))

    # 1. Строки, отличающиеся одной ячейкой, затем двумя. Сравниваем «подписи»
    # строк без одного (двух) столбцов; работает и тогда, когда все значения
    # в строке часто повторяются, а уникального столбца нет.
    ncols = len(norm1[left1[0]])
    h1 = {i: [hash((c, v)) for c, v in enumerate(norm1[i])] for i in left1}
    h2 = {j: [hash((c, v)) for c, v in enumerate(norm2[j])] for j in left2}
    t1 = {i: sum(h) for i, h in h1.items()}
    t2 = {j: sum(h) for j, h in h2.items()}
    drops = [(c,) for c in range(ncols)]
    if ncols * (ncols - 1) // 2 * (len(left1) + len(left2)) <= 5_000_000:
        drops += [(c, d) for c in range(ncols) for d in range(c + 1, ncols)]
    for drop in drops:
        rest1 = [i for i in left1 if i not in used1]
        rest2 = [j for j in left2 if j not in used2]
        if not rest1 or not rest2:
            break
        index = defaultdict(deque)
        for j in rest2:
            index[t2[j] - sum(h2[j][c] for c in drop)].append(j)
        for i in rest1:
            queue = index.get(t1[i] - sum(h1[i][c] for c in drop))
            if queue and ratio(i, queue[0]) >= o.similarity:
                take(i, queue.popleft())

    # 2. По редким совпадающим значениям
    rest1 = [i for i in left1 if i not in used1]
    rest2 = [j for j in left2 if j not in used2]
    candidates = []
    if rest1 and rest2:
        index = defaultdict(list)   # (столбец, значение) -> строки файла 2
        for j in rest2:
            for c, v in enumerate(norm2[j]):
                if v is not None:
                    index[(c, v)].append(j)
        small = len(rest1) * len(rest2) <= 250_000
        for i in rest1:
            if small:               # строк мало: сравниваем каждую с каждой
                options = sorted(rest2, key=lambda j: -same_cells(norm1[i], norm2[j]))[:5]
            else:
                votes = Counter()
                for c, v in enumerate(norm1[i]):
                    if v is not None:
                        bucket = index.get((c, v), ())
                        if len(bucket) <= TOO_COMMON:
                            votes.update(bucket)
                options = [j for j, _ in votes.most_common(5)]
            for j in options:
                r = ratio(i, j)
                if r >= o.similarity:
                    candidates.append((-r, i, j))
    candidates.sort()
    for _, i, j in candidates:
        if i not in used1 and j not in used2:
            take(i, j)
    return pairs


def compare_sheet(res, sheet, df1, df2, o, nm, sheet_ids=None):
    sheet_ids = sheet_ids or {sheet}   # под какими названиями лист известен в --sum
    first_row = o.header + 2   # номер строки Excel для первой строки данных
    A, B, A_in, B_in = nm.a, nm.b, nm.a_in, nm.b_in

    # --- заголовки ---
    def clean_headers(df):
        spaced = [c for c in df.columns if isinstance(c, str) and c != c.strip()]
        stripped = [c.strip() if isinstance(c, str) else c for c in df.columns]
        if len(set(stripped)) == len(stripped):
            df = df.set_axis(stripped, axis=1)
        return df, [c.strip() for c in spaced]

    df1, spaced1 = clean_headers(df1)
    df2, spaced2 = clean_headers(df2)

    # Положение столбцов в исходных файлах (до исключения столбцов из сравнения),
    # чтобы в отчёте указывать букву столбца Excel
    pos1 = {c: k for k, c in enumerate(df1.columns)}
    pos2 = {c: k for k, c in enumerate(df2.columns)}

    orig1, orig2 = df1, df2
    sum_names, sum_absent = [], []
    for sh, k in o.sum_cols:
        if sh is not None and sh not in sheet_ids:
            continue
        if k >= len(df1.columns):
            if sh is None:      # буква без листа: на этом листе такого столбца нет
                sum_absent.append(excel_letter(k))
                continue
            raise ValueError(
                f"Для суммирования указан столбец {excel_letter(k)}, но на листе "
                f"«{sheet}» первого файла столбцы только до "
                f"{excel_letter(len(df1.columns) - 1)}:\n" + column_list(df1.columns))
        res.sum_used.add((sh, k))
        if df1.columns[k] not in sum_names:
            sum_names.append(df1.columns[k])
    if sum_absent:
        res.check(sheet, "Столбцы для суммирования", INFO, comment=
                  f"На этом листе нет {plural(len(sum_absent), 'столбца', 'столбцов', 'столбцов')} "
                  + listing(sum_absent) + ", суммы по ним здесь не считались")

    def cell_col(c):
        """Столбец с буквой Excel: «AF: название» или «AF / AH: название»."""
        letters = list(dict.fromkeys(
            excel_letter(pos[c]) for pos in (pos1, pos2) if c in pos))
        prefix = " / ".join(letters)
        if str(c).startswith("Unnamed: "):
            return f"{prefix} (без заголовка)"
        return f"{prefix}: {c}"
    if spaced1 or spaced2:
        only_one = sorted(set(spaced1) ^ set(spaced2))
        if only_one:
            res.check(sheet, "Заголовки столбцов", WARN, len(spaced1), len(spaced2),
                      "Лишние пробелы в названии только в одном из файлов: "
                      + listing(only_one))
        else:
            res.check(sheet, "Заголовки столбцов", INFO, len(spaced1), len(spaced2),
                      "Лишние пробелы в названиях (одинаково в обоих файлах): "
                      + listing(spaced1))

    # --- какие столбцы сравниваем ---
    skip = set(o.ignore)
    if o.skip_first_column and len(df1.columns):
        # Исключаем столбец A первого файла; во втором файле - столбец с тем же
        # названием, где бы он ни стоял (порядок столбцов может быть разным)
        first1 = df1.columns[0]
        skip.add(first1)
        if len(df2.columns) and df2.columns[0] != first1:
            res.check(sheet, "Столбец A", WARN, comment=
                      f"В файлах он разный: «{display(first1)}» в {A_in} и "
                      f"«{display(df2.columns[0])}» в {B_in}. Исключён только "
                      f"«{display(first1)}»; «{display(df2.columns[0])}» сравнивается")
    skipped = [c for c in dict.fromkeys([*df1.columns, *df2.columns]) if c in skip]
    df1 = df1.drop(columns=[c for c in df1.columns if c in skip])
    df2 = df2.drop(columns=[c for c in df2.columns if c in skip])
    if skipped:
        res.check(sheet, "Исключённые столбцы", INFO, comment="Не сравнивались: "
                  + listing(map(display, skipped)))

    cols1, cols2 = list(df1.columns), list(df2.columns)
    only1 = [c for c in cols1 if c not in cols2]
    only2 = [c for c in cols2 if c not in cols1]
    common = [c for c in cols1 if c in cols2]
    if only1 or only2:
        comment = []
        if only1:
            comment.append(f"только в {A_in}: " + listing(map(display, only1)))
        if only2:
            comment.append(f"только в {B_in}: " + listing(map(display, only2)))
        res.check(sheet, "Набор столбцов", DIFF, len(cols1), len(cols2),
                  cap("; ".join(comment)))
        for c, where, g in [(c, A_in, 1) for c in only1] + [(c, B_in, 2) for c in only2]:
            res.describe(sheet, "Структура", f"Столбец {display(c)}",
                         f"Есть только в {where}")
            res.detail(sheet, g, row_a="все" if g == 1 else None,
                       row_b="все" if g == 2 else None, column=cell_col(c),
                       reason="весь столбец")
            res.totals["structure"] += 1
    else:
        res.check(sheet, "Набор столбцов", OK, len(cols1), len(cols2))
    same_order = common == [c for c in cols2 if c in cols1]
    res.check(sheet, "Порядок столбцов", OK if same_order else INFO,
              comment="" if same_order else
              "Порядок разный; столбцы сопоставлены по названию")

    # --- значения по столбцам ---
    n1, n2 = len(df1), len(df2)
    raw1 = {c: df1[c].tolist() for c in cols1}
    raw2 = {c: df2[c].tolist() for c in cols2}
    nc1 = {c: [canon(v, o) for v in raw1[c]] for c in cols1}
    nc2 = {c: [canon(v, o) for v in raw2[c]] for c in cols2}

    def edge_spaces(raw, cols):
        return sum(1 for c in cols for v in raw[c]
                   if isinstance(v, str) and v and (v[0].isspace() or v[-1].isspace())
                   and v.strip())

    sp1, sp2 = edge_spaces(raw1, cols1), edge_spaces(raw2, cols2)
    if sp1 or sp2:
        res.check(sheet, "Пробелы в начале и конце текста", INFO if sp1 == sp2 else WARN,
                  sp1, sp2, "Сколько ячеек с такими пробелами. При сравнении они не "
                  "учитываются: «Иванов » и «Иванов» считаются одинаковыми")

    # Подсказка про нумерацию в столбце A
    if cols1 and cols2 and cols1[0] == cols2[0] and not o.skip_first_column:
        first = cols1[0]
        unnamed = str(first).startswith("Unnamed: ")
        if unnamed or (looks_like_numbering(nc1[first]) and looks_like_numbering(nc2[first])):
            res.check(sheet, "Столбец A", WARN, comment=
                      "Похож на нумерацию строк и участвует в сравнении. Если порядок "
                      "строк в файлах разный, нумерация даст ложные отличия; "
                      "исключить столбец: --skip-first-column")

    # Столбец с одним значением на весь файл, разным в двух файлах
    # (полностью пустые строки при этом не учитываются)
    def filled_rows(nc, n):
        if not common:
            return []
        return [i for i, r in enumerate(zip(*[nc[c] for c in common]))
                if any(v is not None for v in r)]

    const = {}
    fr1, fr2 = filled_rows(nc1, n1), filled_rows(nc2, n2)
    if len(fr1) > 1 and len(fr2) > 1:
        for c in common:
            u1 = set(nc1[c]) if len(fr1) == n1 else {nc1[c][i] for i in fr1}
            u2 = set(nc2[c]) if len(fr2) == n2 else {nc2[c][i] for i in fr2}
            if len(u1) == 1 and len(u2) == 1 and None not in u1 | u2:
                v1, v2 = raw1[c][fr1[0]], raw2[c][fr2[0]]
                if not values_equal(v1, v2, o):
                    const[c] = (v1, v2)
    # Техническим считаем только отметку времени (дата со временем или время):
    # так выглядит время формирования отчёта. Остальное - расхождение данных
    tech_cols = [c for c, (v1, v2) in const.items() if is_stamp(v1) and is_stamp(v2)]
    whole_cols = [c for c in const if c not in tech_cols]
    for c, (v1, v2) in const.items():
        res.describe(sheet, "Столбец целиком", f"Столбец {display(c)}, все строки",
                     f"В {A_in} везде «{brief(v1, 60)}», в {B_in} везде «{brief(v2, 60)}»")
        res.detail(sheet, 3, row_a="все", row_b="все", column=cell_col(c),
                   val_a=v1, val_b=v2,
                   reason="отличается во всех строках")
        res.totals["const_columns"] += 1
        if c in tech_cols:
            res.tech.append((sheet, cell_col(c), v1, v2))
        else:
            res.whole.append((sheet, cell_col(c), v1, v2))
            res.totals["const_real"] += 1
    res.check(sheet, "Столбцы, отличающиеся во всех строках",
              DIFF if whole_cols else OK, comment=
              listing(map(display, whole_cols)) + ". Одно значение на весь файл, разное "
              "в двух файлах; служебный столбец можно исключить: --ignore"
              if whole_cols else "")
    if tech_cols:
        res.check(sheet, "Технические столбцы", TECH, comment=
                  listing(map(display, tech_cols)) + ". Отметка времени, одна на весь файл")

    cmp_cols = [c for c in common if c not in const]
    rows1 = list(zip(*[raw1[c] for c in cmp_cols])) if cmp_cols else [()] * n1
    rows2 = list(zip(*[raw2[c] for c in cmp_cols])) if cmp_cols else [()] * n2
    norm1 = list(zip(*[nc1[c] for c in cmp_cols])) if cmp_cols else [()] * n1
    norm2 = list(zip(*[nc2[c] for c in cmp_cols])) if cmp_cols else [()] * n2

    def blank(n):
        return all(v is None for v in n)

    active1 = [i for i, n in enumerate(norm1) if not blank(n)]
    active2 = [j for j, n in enumerate(norm2) if not blank(n)]

    # --- общие проверки объёма ---
    res.check(sheet, "Количество строк с данными",
              OK if len(active1) == len(active2) else DIFF, len(active1), len(active2))
    if n1 - len(active1) or n2 - len(active2):
        res.check(sheet, "Полностью пустые строки", INFO,
                  n1 - len(active1), n2 - len(active2), "В сравнении не участвуют")
    empty_cols1 = [c for c in cols1 if n1 and all(v is None for v in nc1[c])]
    empty_cols2 = [c for c in cols2 if n2 and all(v is None for v in nc2[c])]
    if empty_cols1 or empty_cols2:
        same = empty_cols1 == empty_cols2
        res.check(sheet, "Полностью пустые столбцы", INFO if same else WARN,
                  len(empty_cols1), len(empty_cols2),
                  listing(map(display, dict.fromkeys(empty_cols1 + empty_cols2))))
    dup1 = len(active1) - len({norm1[i] for i in active1})
    dup2 = len(active2) - len({norm2[j] for j in active2})
    res.check(sheet, "Повторяющиеся строки", INFO if dup1 == dup2 else WARN, dup1, dup2,
              "Сколько строк повторяют другую строку того же файла" if dup1 or dup2 else "")
    same_pos = sum(a == b for a, b in zip(norm1, norm2))
    if same_pos == n1 == n2:
        res.check(sheet, "Порядок строк", OK)
    else:
        res.check(sheet, "Порядок строк", INFO, comment=
                  f"На тех же позициях стоят {same_pos} из {max(n1, n2)} строк; "
                  "строки сопоставлены независимо от порядка")

    # --- сопоставление строк ---
    candidates, only_a, only_b = [], [], []   # пары (i, j, ключ) и строки без пары
    exact = 0
    method = ""
    keys = [k.strip() for k in (o.keys or [])]
    missing = [k for k in keys if k not in cmp_cols]
    if keys and missing:
        res.check(sheet, "Ключевые столбцы", WARN, comment="Не найдены: "
                  + listing(missing) + ". Строки сопоставлены без ключа")
        keys = []

    if keys:
        method = "по ключевым столбцам: " + listing(keys)
        key_idx = [cmp_cols.index(k) for k in keys]

        def groups_of(norm, active):
            result = defaultdict(list)
            for i in active:
                result[tuple(norm[i][c] for c in key_idx)].append(i)
            return result

        def label(rows, i):
            return ", ".join(f"{k}={rows[i][c]}" for k, c in zip(keys, key_idx))

        g1, g2 = groups_of(norm1, active1), groups_of(norm2, active2)
        repeated = sum(1 for g in (g1, g2) for rows in g.values() if len(rows) > 1)
        for k, part1 in g1.items():
            part2 = g2.get(k, [])
            if len(part1) == 1 and len(part2) == 1:
                candidates.append((part1[0], part2[0], label(rows1, part1[0])))
                continue
            # Ключ повторяется: внутри группы сначала связываем одинаковые строки,
            # затем самые похожие (порядок строк в файлах может быть разным)
            pool = defaultdict(deque)
            for j in part2:
                pool[norm2[j]].append(j)
            rest1 = []
            for i in part1:
                if pool[norm1[i]]:
                    candidates.append((i, pool[norm1[i]].popleft(), label(rows1, i)))
                else:
                    rest1.append(i)
            rest2 = sorted(j for q in pool.values() for j in q)
            if rest1 and rest2:
                if len(rest1) * len(rest2) <= 10_000:
                    order = sorted((-same_cells(norm1[i], norm2[j]), i, j)
                                   for i in rest1 for j in rest2)
                else:
                    order = [(0, i, j) for i, j in zip(rest1, rest2)]
                done1, done2 = set(), set()
                for _, i, j in order:
                    if i not in done1 and j not in done2:
                        done1.add(i)
                        done2.add(j)
                        candidates.append((i, j, label(rows1, i)))
                rest1 = [i for i in rest1 if i not in done1]
                rest2 = [j for j in rest2 if j not in done2]
            only_a += [(i, label(rows1, i)) for i in rest1]
            only_b += [(j, label(rows2, j)) for j in rest2]
        only_b += [(j, label(rows2, j)) for k, part2 in g2.items() if k not in g1
                   for j in part2]
        if repeated:
            res.check(sheet, "Ключевые столбцы", WARN, comment=
                      f"Ключ не уникален: повторяется у {repeated} "
                      f"{plural(repeated, 'значения', 'значений', 'значений')}. Строки "
                      "с одинаковым ключом сопоставлены по сходству")
        total1, total2 = len(active1), len(active2)

    elif o.by_order:
        method = "строго по порядку строк"
        n = min(n1, n2)
        candidates = [(i, i, "") for i in range(n)]
        only_a = [(i, "") for i in range(n, n1)]
        only_b = [(j, "") for j in range(n, n2)]
        total1, total2 = n1, n2

    else:
        # 1. Полностью одинаковые строки (с учётом повторов)
        pool = defaultdict(deque)
        for j in active2:
            pool[norm2[j]].append(j)
        left1 = []
        for i in active1:
            if pool[norm1[i]]:
                pool[norm1[i]].popleft()
                exact += 1
            else:
                left1.append(i)
        left2 = sorted(j for q in pool.values() for j in q)
        method = "по полному совпадению строки"

        # 2. По столбцу-идентификатору
        if left1 and left2:
            kc = find_id_column(cmp_cols, norm1, norm2, active1, active2, left1, left2)
            if kc is not None:
                by_id = {norm2[j][kc]: j for j in left2 if norm2[j][kc] is not None}
                used = set()
                alone = len(cmp_cols) == 1
                for i in left1:
                    j = by_id.get(norm1[i][kc])
                    # Строки связываем, только если кроме идентификатора у них
                    # совпадает что-то ещё: иначе это, скорее, разные записи
                    if (j is not None and j not in used
                            and (alone or same_cells(norm1[i], norm2[j], skip=kc))):
                        used.add(j)
                        candidates.append((i, j, ""))
                if candidates:
                    method += f"; строки с расхождениями - по столбцу {display(cmp_cols[kc])}"
                    e1 = sum(1 for i in active1 if norm1[i][kc] is None)
                    e2 = sum(1 for j in active2 if norm2[j][kc] is None)
                    if e1 or e2:
                        res.check(sheet, "Пустой ключ", INFO, e1, e2,
                                  f"Пустые ячейки в столбце {display(cmp_cols[kc])}; "
                                  "такие строки по нему не связать")
                    done = {i for i, _, _ in candidates}
                    left1 = [i for i in left1 if i not in done]
                    left2 = [j for j in left2 if j not in used]

        # 3. По сходству
        if left1 and left2:
            pairs = pair_by_similarity(left1, left2, norm1, norm2, rows1, rows2, o)
            if pairs:
                method += "; остальные - по сходству"
                candidates += [(i, j, "") for i, j in pairs]
                done1 = {i for i, _ in pairs}
                done2 = {j for _, j in pairs}
                left1 = [i for i in left1 if i not in done1]
                left2 = [j for j in left2 if j not in done2]
        only_a = [(i, "") for i in left1]
        only_b = [(j, "") for j in left2]
        total1, total2 = len(active1), len(active2)

    # --- отличия в сопоставленных строках ---
    const_cols = whole_cols + tech_cols
    table = {"sheet": sheet, "pairs": [], "only": [],
             "columns": [cell_col(c) for c in cmp_cols + const_cols],
             "tech": len(tech_cols), "whole": len(whole_cols)}
    changed = 0
    reasons = Counter()
    col_stats = defaultdict(Counter)
    for i, j, key in sorted(candidates):
        diffs = []
        for c, a, b in zip(cmp_cols, rows1[i], rows2[j]):
            if not values_equal(a, b, o):
                kind, reason = classify(a, b)
                if kind == "empty_a":
                    reason = f"пусто в {A_in}"
                elif kind == "empty_b":
                    reason = f"пусто в {B_in}"
                diffs.append((c, a, b, kind, reason))
        if not diffs:
            exact += 1
            continue
        changed += 1
        ra, rb = i + first_row, j + first_row
        if len(table["pairs"]) < MAX_TABLE_PAIRS:
            hint = []
            for kind, text in (("empty_b", f"{B}: не заполнено"), ("empty_a", f"{A}: не заполнено"),
                               ("value", "разные значения:"), ("form", "отличается форма записи:")):
                names = [display(c) for c, _, _, k, _ in diffs if k == kind]
                if names:
                    hint.append(f"{text} {listing(names, 4)}")
            table["pairs"].append((
                ra, rb, rows1[i] + tuple(raw1[c][i] for c in const_cols),
                rows2[j] + tuple(raw2[c][j] for c in const_cols),
                {cmp_cols.index(c): k for c, _, _, k, _ in diffs}, "; ".join(hint)))
        for c, a, b, kind, reason in diffs:
            reasons[reason] += 1
            col_stats[c]["cells"] += 1
            col_stats[c][kind] += 1
            res.detail(sheet, GROUP_OF[kind], key, ra, rb, cell_col(c), a, b,
                       reason if GROUP_OF[kind] == 3 else "")

        where = (f"Строка {ra} в обоих файлах" if ra == rb
                 else f"Строка {ra} в {A_in}, {rb} в {B_in}")
        if key:
            where = f"{key}: {where[0].lower()}{where[1:]}"
        parts = []
        for kind, side in (("empty_b", B_in), ("empty_a", A_in)):
            names = [display(c) for c, _, _, k, _ in diffs if k == kind]
            if len(names) == 1:
                parts.append(f"в {side} пустое поле {names[0]}")
            elif names:
                parts.append(f"в {side} пустые {len(names)} "
                             f"{plural(len(names), 'поле', 'поля', 'полей')}: "
                             + listing(names))
        real = [(c, a, b) for c, a, b, k, _ in diffs if k == "value"]
        if len(real) == 1:
            c, a, b = real[0]
            parts.append(f"разное значение в поле {display(c)}: в {A_in} "
                         f"«{brief(a)}», в {B_in} «{brief(b)}»")
        elif real:
            parts.append(f"разные значения в {len(real)} полях: "
                         + listing(display(c) for c, _, _ in real))
        form = [f"{display(c)} ({r})" for c, _, _, k, r in diffs if k == "form"]
        if form:
            parts.append(f"отличается только форма записи в {len(form)} "
                         f"{plural(len(form), 'поле', 'полях', 'полях')}: " + listing(form))
        res.describe(sheet, "Строка изменена", where, cap("; ".join(parts)))

    # --- строки без пары ---
    for items, rows, side, other, kind_a in ((only_a, rows1, A_in, B_in, True),
                                             (only_b, rows2, B_in, A_in, False)):
        for i, key in items:
            r = i + first_row
            content = describe_row(cmp_cols, rows[i])
            if len(table["only"]) < MAX_TABLE_PAIRS:
                raw = raw1 if kind_a else raw2
                table["only"].append((A if kind_a else B, r,
                                      rows[i] + tuple(raw[c][i] for c in const_cols),
                                      f"строки нет в {other}"))
            where = f"Строка {r} в {side}" + (f" ({key})" if key else "")
            res.describe(sheet, f"Строка только в {side}", where,
                         f"Строки нет в {other}: {content}")
            res.detail(sheet, 1 if kind_a else 2, key,
                       r if kind_a else None, None if kind_a else r,
                       column="вся строка",
                       val_a=content if kind_a else None,
                       val_b=None if kind_a else content)

    # --- итоги по строкам ---
    res.check(sheet, "Способ сопоставления строк", INFO, comment=cap(method))
    res.check(sheet, "Строки, совпавшие полностью", INFO, exact, exact)
    res.check(sheet, "Строки с отличиями в ячейках", DIFF if changed else OK,
              changed, changed)
    if reasons:
        res.check(sheet, "Отличия в ячейках по причинам", INFO, comment="; ".join(
            f"{r}: {n}" for r, n in reasons.most_common()))
    res.check(sheet, "Строки только в одном из файлов",
              DIFF if only_a or only_b else OK, len(only_a), len(only_b),
              f"Слева - строки, которых нет в {B_in}; справа - которых нет в {A_in}"
              if only_a or only_b else "")
    cells = Counter()
    for st in col_stats.values():
        cells[1] += st["empty_b"]
        cells[2] += st["empty_a"]
        cells[3] += st["value"] + st["form"]
    for g, title, rows_n, cols_n in (
            (1, f"Есть в {A_in}, нет в {B_in}", len(only_a), len(only1)),
            (2, f"Нет в {A_in}, есть в {B_in}", len(only_b), len(only2)),
            (3, "Значения разные", 0, 0)):
        parts = [f"ячеек: {cells[g]}"]
        if g != 3:
            parts.append(f"строк целиком: {rows_n}")
        if cols_n:
            parts.append(f"столбцов целиком: {cols_n}")
        res.check(sheet, title, DIFF if cells[g] or rows_n or cols_n else OK,
                  comment=cap("; ".join(parts)))
    balanced = (exact + changed + len(only_a) == total1
                and exact + changed + len(only_b) == total2)
    res.check(sheet, "Баланс строк", OK if balanced else WARN, total1, total2,
              f"совпало {exact} + с расхождениями {changed} + только в одном файле "
              + (f"{len(only_a)} = {total1}" if len(only_a) == len(only_b) and total1 == total2
                 else f"{len(only_a)} / {len(only_b)}")
              + ("" if balanced else ". Баланс не сходится, результат требует проверки"))

    # --- суммы по столбцам, указанным в --sum ---
    sums, sum_diff = {}, []
    for c in sum_names:
        sa, bad_a = column_sum(orig1[c].tolist())
        sb, bad_b = column_sum(orig2[c].tolist()) if c in orig2.columns else (None, 0)
        differs = sb is None or abs(sa - sb) > max(o.tolerance, 1e-6)
        sums[c] = (sa, sb, bad_a, bad_b, differs)
        if differs:
            sum_diff.append(c)
    if sum_names:
        res.check(sheet, "Суммы по указанным столбцам", DIFF if sum_diff else OK,
                  comment=(f"Расходятся по {len(sum_diff)} из {len(sum_names)}: "
                           + listing(map(cell_col, sum_diff))) if sum_diff
                  else f"Сходятся по всем указанным столбцам ({len(sum_names)})")

    # --- профиль столбцов ---
    type_diff = []
    for c in dict.fromkeys(cols1 + cols2 + sum_names):
        entry = {"Лист": sheet, "Столбец": cell_col(c)}
        notes, types = [], {}
        for side, raw, nc in (("a", raw1, nc1), ("b", raw2, nc2)):
            if c not in raw:
                continue
            filled = [v for v in raw[c] if not is_empty(v)]
            entry[f"filled_{side}"] = len(filled)
            entry[f"empty_{side}"] = len(raw[c]) - len(filled)
            entry[f"unique_{side}"] = len(set(nc[c]) - {None})
            types[side] = sorted({type_label(v) for v in filled})
        if c in sums:
            sa, sb, bad_a, bad_b, differs = sums[c]
            entry["sum"] = sa if not differs else f"{number_text(sa)} / {number_text(sb)}"
            entry["sum_a"], entry["sum_b"], entry["sum_differs"] = sa, sb, differs
            if differs:
                notes.append("суммы расходятся")
            for bad, side in ((bad_a, A_in), (bad_b, B_in)):
                if bad:
                    notes.append(f"в {side} {bad} "
                                 f"{plural(bad, 'ячейка', 'ячейки', 'ячеек')} не "
                                 f"{plural(bad, 'число', 'числа', 'числа')}, в сумму не "
                                 f"{plural(bad, 'вошла', 'вошли', 'вошли')}")
        if c not in cols1 and c not in cols2:
            notes.append("столбец исключён из сравнения")
        if len(types) == 2 and types["a"] != types["b"] and types["a"] and types["b"]:
            type_diff.append(c)
            notes.append(f"тип данных отличается: {', '.join(types['a'])} / "
                         f"{', '.join(types['b'])}")
        elif any(len(t) > 1 for t in types.values()):
            notes.append("в столбце смешаны разные типы данных")
        if c in const:
            notes.append("во всех строках разное значение")
        if c in spaced1 or c in spaced2:
            notes.append("лишние пробелы в заголовке")
        st = col_stats.get(c, {})
        entry["g1"] = "весь столбец" if c in only1 else st.get("empty_b", 0)
        entry["g2"] = "весь столбец" if c in only2 else st.get("empty_a", 0)
        entry["g3"] = "все" if c in const else st.get("value", 0) + st.get("form", 0)
        entry["Замечания"] = "; ".join(notes)
        res.columns.append(entry)

    if type_diff:
        res.check(sheet, "Типы данных в столбцах", WARN, comment=
                  "Тип отличается между файлами (например, число и текст): "
                  + listing(map(display, type_diff)))
    else:
        res.check(sheet, "Типы данных в столбцах", OK)

    table["changed"], table["only_n"] = changed, len(only_a) + len(only_b)
    res.tables.append(table)
    res.totals["exact"] += exact
    for g in (1, 2, 3):
        res.totals[f"cells{g}"] += cells[g]
    res.totals["rows_a"] += len(active1)
    res.totals["rows_b"] += len(active2)
    res.totals["changed"] += changed
    res.totals["only_a"] += len(only_a)
    res.totals["only_b"] += len(only_b)


# ---------- сравнение файлов ----------

# Текст «NA», «NULL», «None», «n/a» и подобный читаем как текст, а не как пустую ячейку
READ_OPTIONS = {"keep_default_na": False, "na_filter": False}
FAST_READER = None   # определяется при первом чтении: True - calamine, False - openpyxl


def read_book(path, header):
    """Читает все листы файла. По возможности использует быструю библиотеку calamine."""
    global FAST_READER
    if FAST_READER is None:
        try:
            import python_calamine  # noqa: F401
            FAST_READER = True
        except ImportError:
            FAST_READER = False
    if FAST_READER:
        try:
            return pd.read_excel(path, sheet_name=None, header=header, dtype=object,
                                 engine="calamine", **READ_OPTIONS)
        except Exception:
            # Старая версия pandas или файл, который calamine не смог прочитать:
            # оба файла читаем обычным способом, чтобы типы значений совпадали
            FAST_READER = False
            raise RereadNeeded
    return pd.read_excel(path, sheet_name=None, header=header, dtype=object,
                         engine="openpyxl", **READ_OPTIONS)


class RereadNeeded(Exception):
    pass


def compare_files(path1, path2, o, nm):
    try:
        book1 = read_book(path1, o.header)
        book2 = read_book(path2, o.header)
    except RereadNeeded:
        book1 = read_book(path1, o.header)
        book2 = read_book(path2, o.header)
    res = Result()
    res.books = (book1, book2)

    # Если в каждом файле один лист, сравниваем их даже при разных названиях
    if len(book1) == 1 and len(book2) == 1 and list(book1) != list(book2) and not o.sheets:
        (s1, d1), (s2, d2) = *book1.items(), *book2.items()
        res.check("", "Набор листов", INFO, 1, 1,
                  f"Единственные листы названы по-разному: «{s1}» и «{s2}»")
        compare_sheet(res, f"{s1} / {s2}", d1, d2, o, nm, sheet_ids={s1, s2})
        check_sum_columns(res, o, book1)
        return res

    names = o.sheets or list(dict.fromkeys([*book1, *book2]))
    missing = [n for n in names if n not in book1 and n not in book2]
    if missing:
        raise ValueError("Лист не найден ни в одном файле: " + listing(missing))
    only1 = [n for n in names if n not in book2]
    only2 = [n for n in names if n not in book1]
    if only1 or only2:
        comment = []
        if only1:
            comment.append(f"только в {nm.a_in}: " + listing(only1))
        if only2:
            comment.append(f"только в {nm.b_in}: " + listing(only2))
        res.check("", "Набор листов", DIFF, len(book1), len(book2), cap("; ".join(comment)))
        for n, where, g in [(n, nm.a_in, 1) for n in only1] + [(n, nm.b_in, 2) for n in only2]:
            res.describe(n, "Структура", f"Лист {n}", f"Есть только в {where}")
            res.detail(n, g, column=f"весь лист «{n}»", reason="весь лист")
            res.totals["structure"] += 1
    else:
        res.check("", "Набор листов", OK, len(book1), len(book2))

    for n in names:
        if n in book1 and n in book2:
            compare_sheet(res, n, book1[n], book2[n], o, nm)
    check_sum_columns(res, o, book1)
    return res


def check_sum_columns(res, o, book1):
    """Ошибка, если столбец из --sum не нашёлся ни на одном сравниваемом листе."""
    for sh, k in o.sum_cols:
        if (sh, k) in res.sum_used:
            continue
        if sh is not None:
            raise ValueError(f"Для суммирования указан лист «{sh}», но такого листа нет "
                             "среди сравниваемых: " + listing(book1))
        name, df = next(iter(book1.items()))
        raise ValueError(
            f"Для суммирования указан столбец {excel_letter(k)}, но такого столбца нет "
            f"ни на одном листе первого файла. Столбцы листа «{name}»:\n"
            + column_list(df.columns))


# ---------- отчёт ----------

FILLS = {OK: "C6EFCE", DIFF: "FFC7CE", WARN: "FFEB9C", INFO: "EDEDED"}
HEADER_FILL = PatternFill("solid", fgColor="305496")


def clean(v):
    """Значение, пригодное для записи в Excel."""
    if is_empty(v):
        return None
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)[:32000]
        return "'" + v if v.startswith("=") else v
    return v


def style_table(ws, header_row, wrap=()):
    for cell in ws[header_row]:
        if cell.value is not None:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for col in ws.iter_cols(min_row=header_row, max_row=min(ws.max_row, header_row + 300)):
        letter = col[0].column_letter
        width = max((len(str(c.value)) for c in col if c.value is not None), default=0)
        limit = 90 if letter in wrap else 45
        ws.column_dimensions[letter].width = max(10, min(width + 2, limit))
    for letter in wrap:
        for cell in ws[letter][header_row:min(ws.max_row, 20000)]:
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def color_rows(ws, column, last_col, rules):
    """Подсветка строк по значению в столбце (условное форматирование)."""
    if ws.max_row < 2:
        return
    area = f"A2:{last_col}{ws.max_row}"
    for text, color in rules:
        quoted = text.replace('"', '""')
        ws.conditional_formatting.add(area, FormulaRule(
            formula=[f'${column}2="{quoted}"'],
            fill=PatternFill("solid", bgColor=color, fgColor=color)))


THIN = Side(style="thin", color="BFBFBF")
MEDIUM = Side(style="medium", color="7F7F7F")
ROW_FORMAT = r"#\ ##0"


def put(ws, row, col, value, *, bold=False, italic=False, color=None, fill=None, size=10,
        align=None, valign="top", wrap=True, border=True, fmt=None, heavy=False):
    """Записывает ячейку листа «Итог» с оформлением."""
    cell = ws.cell(row, col, clean(value))
    cell.font = Font(bold=bold, italic=italic, size=size, color=color)
    cell.alignment = Alignment(horizontal=align, vertical=valign, wrap_text=wrap)
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    if border:
        cell.border = Border(left=THIN, right=THIN, top=THIN,
                             bottom=MEDIUM if heavy else THIN)
    if fmt:
        cell.number_format = fmt
    elif isinstance(cell.value, str):
        cell.number_format = "@"
    return cell


def wide(ws, row, first, last, heavy=False):
    """Объединяет ячейки строки и рисует рамку по всей ширине."""
    if last > first:
        for col in range(first + 1, last + 1):
            ws.cell(row, col).border = Border(left=THIN, right=THIN, top=THIN,
                                              bottom=MEDIUM if heavy else THIN)
        ws.merge_cells(start_row=row, start_column=first, end_row=row, end_column=last)


def tech_comment(v1, v2):
    """Пояснение к столбцу с одним значением на файл."""
    tail = "Отличается во всех строках, расхождением данных не считается."
    try:
        t1, t2 = (pd.to_datetime(v, errors="raise") for v in (v1, v2))
        if not is_number(v1) and not is_number(v2):
            seconds = abs((t1 - t2).total_seconds())
            gap = (f"{seconds:.0f} с" if seconds < 600 else f"{seconds / 60:.0f} мин"
                   if seconds < 36000 else f"{seconds / 3600:.1f} ч")
            return ("Похоже на время формирования отчёта. Одно значение на весь файл; "
                    f"разница между файлами {gap}. {tail}")
    except Exception:
        pass
    return f"Одно значение на весь файл, разное в двух файлах. {tail}"


def build_summary_sheet(ws, res, file1, file2, nm, extra_tabs):
    """Лист «Итог»: вывод, таблица расхождений, технические отличия, проверки."""
    A, B = nm.a, nm.b
    t = res.totals
    ws.sheet_view.showGridLines = False
    span = 9   # ширина объединённых строк (A:I)
    many_sheets = len(res.tables) > 1

    def title(row, text, height=21.75):
        cell = put(ws, row, 1, text, bold=True, size=11, border=False, wrap=False,
                   valign="center")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
        ws.row_dimensions[row].height = height
        return cell

    def header(row, labels, merge_last_to=None):
        for col, label in enumerate(labels, 1):
            put(ws, row, col, label, bold=True, color="FFFFFF", fill="305496",
                align="center", valign="center")
        if merge_last_to:
            wide(ws, row, len(labels), merge_last_to)

    # --- шапка и вывод ---
    name = f"Сверка {nm.prefix}: " if nm.prefix else "Сверка: "
    put(ws, 1, 1, f"{name}{A} ↔ {B}", bold=True, size=14, border=False, wrap=False,
        valign="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    ws.row_dimensions[1].height = 25.5
    put(ws, 2, 1, f"{A}: {Path(file1).name}   ·   {B}: {Path(file2).name}   ·   "
        f"дата сверки: {datetime.now():%Y-%m-%d %H:%M:%S}", border=False, wrap=False,
        color="595959")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)

    cells_total = t["cells1"] + t["cells2"] + t["cells3"]
    only_total = t["only_a"] + t["only_b"]
    share = t["exact"] / t["rows_a"] if t["rows_a"] else 1.0
    if res.has_differences:
        parts = ["ИТОГ: РАСХОЖДЕНИЯ В ДАННЫХ"]
        if t["changed"]:
            parts.append(f"строк: {t['changed']} из {t['rows_a']}")
            parts.append(f"ячеек: {cells_total}")
        if only_total:
            parts.append(f"строк только в одном файле: {only_total}")
        if t["structure"]:
            parts.append(f"отличий в структуре: {t['structure']}")
        if t["const_real"]:
            parts.append(f"столбцов отличается во всех строках: {t['const_real']}")
        else:
            parts.append(f"совпадение {round(share * 100, 3):g}%")
        verdict, fill, color = "   ·   ".join(parts), "FFC7CE", "9C0006"
    else:
        verdict, fill, color = "ИТОГ: ДАННЫЕ СОВПАДАЮТ", "C6EFCE", "006100"
    put(ws, 4, 1, verdict, bold=True, size=13, color=color, fill=fill, border=False,
        wrap=False, valign="center")
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=span)
    ws.row_dimensions[4].height = 30
    nature = []
    if cells_total:
        nature.append(f"пусто в {nm.a_in} — {t['cells2']} яч.   ·   пусто в {nm.b_in} — "
                      f"{t['cells1']} яч.   ·   значения разные — {t['cells3']} яч.")
    if res.tech:
        nature.append("технические отличия исключены (см. ниже)")
    if nature:
        put(ws, 5, 1, "Характер расхождений:   " + "   ·   ".join(nature), fill="FFF2CC",
            border=False, wrap=False, valign="center")
        ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=span)
        ws.row_dimensions[5].height = 21.75

    figures = [(f"Строк в {nm.a_in}", t["rows_a"], None),
               (f"Строк в {nm.b_in}", t["rows_b"], None),
               ("Совпали полностью", t["exact"], share),
               ("Строк с расхождениями", t["changed"],
                t["changed"] / t["rows_a"] if t["rows_a"] else 0),
               ("Ячеек с расхождениями", cells_total, None),
               ("Строк только в одном из файлов", only_total, None)]
    for row, (label, value, part) in enumerate(figures, 7):
        put(ws, row, 1, label, fill="EDEDED")
        put(ws, row, 2, value, bold=True, align="right", wrap=False, fmt=ROW_FORMAT)
        if part is not None:
            put(ws, row, 3, part, border=False, wrap=False, fmt="0.000%", color="595959")
    row = 14

    # --- таблицы расхождений ---
    widths = {}

    def data_cell(r, col, value, kind, tech, heavy):
        if tech:
            put(ws, r, col, value, color="7F7F7F", fill="EDEDED", heavy=heavy)
        elif kind is None:
            put(ws, r, col, value, color="A6A6A6", italic=True, heavy=heavy)
        elif is_empty(value):
            put(ws, r, col, "(пусто)", bold=True, italic=True, color="9C0006",
                fill="FFC7CE", heavy=heavy)
        else:
            put(ws, r, col, value, bold=True, heavy=heavy,
                fill="FFEB9C" if kind in ("empty_a", "empty_b") else "FFD966")
        widths[col] = max(widths.get(col, 0), min(len(str(value)), 40))

    truncated = False
    for table in res.tables:
        if not table["pairs"] and not table["only"]:
            continue
        columns, ntech, nwhole = table["columns"], table["tech"], table["whole"]
        ncol = len(columns)
        first_tech = ncol - ntech
        first_whole = first_tech - nwhole
        col_n, col_hint = 3 + ncol, 4 + ncol
        labels = (["Система", "№ строки\nв файле"]
                  + [c + ("\n(технический, исключён)" if k >= first_tech else
                          "\n(отличается во всех строках)" if k >= first_whole else "")
                     for k, c in enumerate(columns)] + ["Отличий,\nячеек", "Что проверить"])
        for k, c in enumerate(columns):
            widths[3 + k] = max(widths.get(3 + k, 0), min(len(c), 28))

        if table["pairs"]:
            sheet_note = f" (лист «{table['sheet']}»)" if many_sheets else ""
            title(row, "Расхождения в данных — строки из обоих файлов, отличающиеся "
                       "ячейки подсвечены" + sheet_note)
            header(row + 1, labels)
            ws.row_dimensions[row + 1].height = 43.5
            row += 2
            for ra, rb, va, vb, kinds, hint in table["pairs"]:
                for r, system, number, values, heavy in ((row, A, ra, va, False),
                                                         (row + 1, B, rb, vb, True)):
                    put(ws, r, 1, system, bold=True, fill="EDEDED", heavy=heavy)
                    put(ws, r, 2, number, fmt=ROW_FORMAT, heavy=heavy)
                    for k, value in enumerate(values):
                        kind = kinds.get(k, "value" if first_whole <= k < first_tech else None)
                        data_cell(r, 3 + k, value, kind, k >= first_tech, heavy)
                put(ws, row, col_n, len(kinds) + nwhole, bold=True, align="center", valign="center",
                    wrap=False)
                put(ws, row + 1, col_n, None, heavy=True)
                ws.merge_cells(start_row=row, start_column=col_n, end_row=row + 1,
                               end_column=col_n)
                put(ws, row, col_hint, hint, valign="center")
                put(ws, row + 1, col_hint, None, heavy=True)
                ws.merge_cells(start_row=row, start_column=col_hint, end_row=row + 1,
                               end_column=col_hint)
                row += 2
            if table["changed"] > len(table["pairs"]):
                truncated = True
                put(ws, row, 1, f"Показаны первые {len(table['pairs'])} записей из "
                    f"{table['changed']}. Все расхождения - на листе «Различия по ячейкам».",
                    italic=True, color="9C0006", border=False, wrap=False)
                row += 1
            # обозначения
            put(ws, row, 1, "Обозначения:", bold=True, border=False)
            put(ws, row, 2, "(пусто)", bold=True, italic=True, color="9C0006", fill="FFC7CE")
            put(ws, row, 3, "значения нет в этой системе", border=False)
            put(ws, row, 4, "значение", bold=True, fill="FFEB9C")
            put(ws, row, 5, "есть только в этой системе", border=False)
            put(ws, row, 6, "значение", bold=True, fill="FFD966")
            put(ws, row, 7, "в системах разное", border=False)
            put(ws, row, 8, "значение", italic=True, color="A6A6A6")
            put(ws, row, 9, "совпадает", border=False)
            if ntech:
                put(ws, row, 10, "значение", color="7F7F7F", fill="EDEDED")
                put(ws, row, 11, "технический столбец, в итог не входит", border=False)
            ws.row_dimensions[row].height = 23.85
            row += 2

        if table["only"]:
            sheet_note = f" (лист «{table['sheet']}»)" if many_sheets else ""
            title(row, "Строки только в одном из файлов" + sheet_note)
            header(row + 1, labels[:2 + ncol] + ["", "Что проверить"])
            ws.row_dimensions[row + 1].height = 43.5
            row += 2
            for system, number, values, hint in table["only"]:
                fill = "F8CBAD" if system == A else "BDD7EE"
                put(ws, row, 1, system, bold=True, fill="EDEDED")
                put(ws, row, 2, number, fmt=ROW_FORMAT)
                for k, value in enumerate(values):
                    if k >= first_tech:
                        put(ws, row, 3 + k, value, color="7F7F7F", fill="EDEDED")
                    else:
                        put(ws, row, 3 + k, value, fill=fill)
                    widths[3 + k] = max(widths.get(3 + k, 0), min(len(str(value)), 40))
                put(ws, row, col_n, None)
                put(ws, row, col_hint, hint)
                row += 1
            if table["only_n"] > len(table["only"]):
                truncated = True
                put(ws, row, 1, f"Показаны первые {len(table['only'])} строк из "
                    f"{table['only_n']}. Все строки - на листе «Различия по ячейкам».",
                    italic=True, color="9C0006", border=False, wrap=False)
                row += 1
            row += 1
        widths[col_n] = max(widths.get(col_n, 0), 10)
        widths[col_hint] = max(widths.get(col_hint, 0), 44)

    # --- отличия в структуре ---
    structure = [d for d in res.descriptions if d["Категория"] == "Структура"]
    if structure:
        title(row, "Отличия в структуре файлов")
        header(row + 1, ["Где", "Что отличается"], merge_last_to=span)
        row += 2
        for d in structure:
            where = d["Где"] + (f" (лист «{d['Лист']}»)" if many_sheets and
                                d["Где"].startswith("Столбец") else "")
            put(ws, row, 1, where)
            put(ws, row, 2, d["Что отличается"], fill="FFC7CE")
            wide(ws, row, 2, span)
            row += 1
        row += 1

    # --- суммы по столбцам (--sum) ---
    sums = [e for e in res.columns if "sum_a" in e]
    if sums:
        title(row, "Суммы по указанным столбцам")
        header(row + 1, ["Столбец", A, B, "Результат"], merge_last_to=span)
        row += 2
        for e in sums:
            bad = e["sum_differs"]
            put(ws, row, 1, e["Столбец"] + (f" (лист «{e['Лист']}»)" if many_sheets else ""))
            for col, value in ((2, e["sum_a"]), (3, e["sum_b"])):
                put(ws, row, col, "нет столбца" if value is None else value, align="right",
                    wrap=False, fmt=ROW_FORMAT if isinstance(value, int) else "#,##0.00####")
            put(ws, row, 4, ("✗ суммы расходятся" if bad else "✓ суммы сходятся")
                + (f". {cap(e['Замечания'])}" if "не числ" in e["Замечания"] else ""),
                fill="FFC7CE" if bad else "C6EFCE", color="9C0006" if bad else "006100")
            wide(ws, row, 4, span)
            row += 1
        row += 1

    # --- столбцы, отличающиеся во всех строках ---
    if res.whole:
        title(row, "Столбцы, отличающиеся во всех строках — расхождение "
                   "(в подсчёт строк и ячеек выше не входят)")
        header(row + 1, ["Столбец", A, B, "Комментарий"], merge_last_to=span)
        row += 2
        for sheet, label, v1, v2 in res.whole:
            put(ws, row, 1, label + (f" (лист «{sheet}»)" if many_sheets else ""))
            put(ws, row, 2, v1, fill="FFC7CE")
            put(ws, row, 3, v2, fill="FFC7CE")
            put(ws, row, 4, "Одно значение на весь файл, разное в двух файлах. Если столбец "
                "служебный, его можно исключить из сверки: --ignore")
            wide(ws, row, 4, span)
            ws.row_dimensions[row].height = 27
            row += 1
        row += 1

    # --- технические отличия ---
    if res.tech:
        title(row, "Исключено из сверки — технические отличия (на итог не влияют)")
        header(row + 1, ["Столбец", A, B, "Комментарий"], merge_last_to=span)
        row += 2
        for sheet, label, v1, v2 in res.tech:
            put(ws, row, 1, label + (f" (лист «{sheet}»)" if many_sheets else ""))
            put(ws, row, 2, v1)
            put(ws, row, 3, v2)
            put(ws, row, 4, tech_comment(v1, v2))
            wide(ws, row, 4, span)
            ws.row_dimensions[row].height = 27
            row += 1
        row += 1

    # --- проверки ---
    shown_elsewhere = {"Столбцы, отличающиеся во всех строках", "Технические столбцы",
                       "Строки, совпавшие полностью",
                       "Строки с отличиями в ячейках", "Отличия в ячейках по причинам",
                       "Суммы по указанным столбцам", "Значения разные",
                       f"Есть в {nm.a_in}, нет в {nm.b_in}", f"Нет в {nm.a_in}, есть в {nm.b_in}"}
    ok_comments = {"Набор листов": "листов в каждом файле",
                   "Набор столбцов": "столбцов в каждом файле",
                   "Количество строк с данными": "строк в каждом файле",
                   "Повторяющиеся строки": "в каждом файле"}
    remarks, passed, notes = [], [], []
    for c in res.checks:
        if c["Проверка"] in shown_elsewhere:
            continue
        label = (f"[{c['Лист']}] " if many_sheets and c["Лист"] else "") + c["Проверка"]
        same = c["a"] == c["b"]
        value = c["a"] if same else f"{number_text(c['a']) if is_number(c['a']) else c['a']} / " \
                                    f"{number_text(c['b']) if is_number(c['b']) else c['b']}"
        status = c["Результат"]
        if status == INFO and c["Проверка"] == "Повторяющиеся строки" and same:
            status = OK
        if status == OK:
            passed.append((label, value, ok_comments.get(c["Проверка"], c["Комментарий"])))
        elif status == INFO:
            text = c["Комментарий"]
            if value != "":
                count = (f"в каждом файле: {number_text(value) if is_number(value) else value}"
                         if same else f"{A} / {B}: {value}")
                text = f"{text} ({count})" if text else cap(count)
            notes.append((label, text))
        else:
            remarks.append((label, status, value, c["Комментарий"]))

    if remarks:
        title(row, "Замечания")
        header(row + 1, ["Проверка", "Результат", "Значение", "Комментарий"],
               merge_last_to=span)
        row += 2
        for label, status, value, comment in remarks:
            bad = status == DIFF
            put(ws, row, 1, label)
            put(ws, row, 2, ("✗ " if bad else "! ") + status, fill="FFC7CE" if bad else "FFEB9C",
                color="9C0006" if bad else "7F6000")
            put(ws, row, 3, value, fmt=ROW_FORMAT if is_number(value) else None)
            put(ws, row, 4, comment)
            wide(ws, row, 4, span)
            ws.row_dimensions[row].height = 15 * max(1, -(-len(str(comment)) // 95))
            row += 1
        row += 1
    if passed:
        title(row, "Проверки без замечаний")
        header(row + 1, ["Проверка", "Результат", "Значение", "Комментарий"],
               merge_last_to=span)
        row += 2
        for label, value, comment in passed:
            put(ws, row, 1, label)
            put(ws, row, 2, "✓ совпадает", fill="C6EFCE", color="006100")
            put(ws, row, 3, value, fmt=ROW_FORMAT if is_number(value) else None)
            put(ws, row, 4, comment)
            wide(ws, row, 4, span)
            row += 1
        row += 1
    if notes:
        title(row, "Справочно — на итог не влияет")
        row += 1
        for label, text in notes:
            put(ws, row, 1, label, fill="EDEDED")
            put(ws, row, 2, text)
            wide(ws, row, 2, span)
            ws.row_dimensions[row].height = 15 * max(1, -(-len(str(text)) // 130))
            row += 1
        row += 1
    if extra_tabs:
        put(ws, row, 1, "Подробности - на листах «Различия по ячейкам», «Столбцы», "
            "«Описание различий», «Все проверки».", italic=True, color="7F7F7F",
            border=False, wrap=False)

    # --- ширина столбцов ---
    base = {1: 32, 2: 13, 3: 24, 4: 22}
    last = max([span] + list(widths))
    for col in range(1, last + 1):
        width = max(base.get(col, 12), widths.get(col, 0) + 2)
        ws.column_dimensions[excel_letter(col - 1)].width = min(width, 46)
    ws.freeze_panes = "C1"
    return truncated


def write_report(path, res, file1, file2, nm, o):
    """Отчёт: лист «Итог» и, при --full или большом числе расхождений, подробные листы."""
    overflow = any(tb["changed"] > MAX_TABLE_PAIRS or tb["only_n"] > MAX_TABLE_PAIRS
                   for tb in res.tables)
    extra = o.full or overflow
    if extra:
        write_detail_tabs(path, res, file1, file2, nm, o)
        book = load_workbook(path)
        book["Сводка"].title = "Все проверки"
        ws = book.create_sheet("Итог", 0)
    else:
        book = Workbook()
        ws = book.active
        ws.title = "Итог"
    build_summary_sheet(ws, res, file1, file2, nm, extra)
    book.active = 0
    for sheet in book.worksheets:
        sheet.sheet_view.tabSelected = sheet is ws
    book.save(path)


def write_detail_tabs(path, res, file1, file2, nm, o):
    A, B = nm.a, nm.b
    truncated = []

    def limited(rows, title):
        if len(rows) > MAX_REPORT_ROWS:
            truncated.append(f"{title}: показаны первые {MAX_REPORT_ROWS} из {len(rows)}")
            return rows[:MAX_REPORT_ROWS]
        return rows

    verdict = "ЕСТЬ РАЗЛИЧИЯ" if res.has_differences else "ФАЙЛЫ СОВПАДАЮТ"
    t = res.totals
    info = [
        (A, str(file1)), (B, str(file2)),
        ("Дата сверки", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Итог", verdict),
        ("Коротко", f"строк: {t['rows_a']} / {t['rows_b']}; изменено строк: {t['changed']}; "
                    f"только в {nm.a_in}: {t['only_a']}; только в {nm.b_in}: {t['only_b']}; "
                    f"столбцов отличается целиком: {t['const_columns']}; "
                    f"отличий в структуре: {t['structure']}"),
    ]
    checks = pd.DataFrame(res.checks).rename(columns={"a": A, "b": B})
    desc = pd.DataFrame(limited(res.descriptions, "Описание различий"),
                        columns=["Лист", "Категория", "Где", "Что отличается"])
    if desc.empty:
        desc.loc[0] = ["", "", "", "Различий не найдено"]
    groups = {1: f"Есть в {nm.a_in}, нет в {nm.b_in}",
              2: f"Нет в {nm.a_in}, есть в {nm.b_in}", 3: "Значения разные"}
    group_colors = {1: "F8CBAD", 2: "BDD7EE", 3: "FFF2CC"}
    many_sheets = len({x["Лист"] for x in res.checks + res.details if x["Лист"]}) > 1

    # Различия по ячейкам: сначала потерянные данные, затем лишние, затем разные
    details = sorted(res.details, key=lambda d: d["group"])   # сортировка устойчивая
    det = pd.DataFrame(limited(details, "Различия по ячейкам"),
                       columns=["group", "Лист", "Ключ", "row_a", "row_b", "Столбец",
                                "val_a", "val_b", "Причина"])
    det["group"] = det["group"].map(groups)
    if (det["Ключ"] == "").all():
        det = det.drop(columns="Ключ")
    if not many_sheets:
        det = det.drop(columns="Лист")
    det = det.rename(columns={"group": "Что с данными",
                              "row_a": f"Строка ({A})", "row_b": f"Строка ({B})",
                              "val_a": f"Значение ({A})", "val_b": f"Значение ({B})"})

    # Столбцы: одно число, если в обоих файлах оно одинаковое, иначе «a / b»
    def num(v):
        return "нет" if v is None else f"{v:,}".replace(",", " ")

    def pair(e, name):
        a, b = e.get(f"{name}_a"), e.get(f"{name}_b")
        return a if a == b else f"{num(a)} / {num(b)}"

    pair_titles = {"filled": "Заполнено", "empty": "Пустых", "unique": "Разных значений"}
    with_sums = any("sum" in e for e in res.columns)
    col_rows = []
    for e in res.columns:
        row = {"Лист": e["Лист"], "Столбец": e["Столбец"]}
        for name, title in pair_titles.items():
            row[f"{title}\n({A} / {B})"] = pair(e, name)
        if with_sums:
            row[f"Сумма\n({A} / {B})"] = e.get("sum")
        for g in (1, 2, 3):
            row[groups[g]] = e[f"g{g}"]
        row["Замечания"] = e["Замечания"]
        problem = (any(isinstance(v, str) for k, v in row.items() if "\n" in k)
                   or any(e[f"g{g}"] for g in (1, 2, 3)))
        col_rows.append((not problem, row))
    col_rows.sort(key=lambda x: x[0])   # столбцы с расхождениями идут первыми
    cols = pd.DataFrame([r for _, r in col_rows])
    if not many_sheets and "Лист" in cols:
        cols = cols.drop(columns="Лист")
    if truncated:
        info.append(("Внимание", "; ".join(truncated)))

    # Любой текст из исходных файлов может содержать символы, которые Excel
    # не принимает, поэтому чистим все таблицы отчёта целиком
    def safe(df):
        return df.apply(lambda col: col.map(lambda v: clean(v) if isinstance(v, str) else v))

    checks, desc, det, cols = safe(checks), safe(desc), safe(det), safe(cols)
    info = [(k, clean(v)) for k, v in info]

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(info).to_excel(writer, sheet_name="Сводка", index=False, header=False)
        start = len(info) + 1
        checks.to_excel(writer, sheet_name="Сводка", index=False, startrow=start)
        desc.to_excel(writer, sheet_name="Описание различий", index=False)
        det.to_excel(writer, sheet_name="Различия по ячейкам", index=False)
        cols.to_excel(writer, sheet_name="Столбцы", index=False)

        ws = writer.sheets["Сводка"]
        style_table(ws, start + 1, wrap=("F",))
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 42
        for row in ws.iter_rows(min_row=1, max_row=len(info)):
            row[0].font = Font(bold=True)
        ws.cell(4, 2).font = Font(bold=True)
        ws.cell(4, 2).fill = PatternFill(
            "solid", fgColor=FILLS[DIFF if res.has_differences else OK])
        for row in ws.iter_rows(min_row=start + 2, max_row=ws.max_row):
            if row[2].value in FILLS:
                row[2].fill = PatternFill("solid", fgColor=FILLS[row[2].value])
        ws.freeze_panes = ws.cell(start + 2, 1)

        row_rules = [("Структура", "F4B084"), ("Столбец целиком", "FFE699"),
                     ("Строка изменена", "FFF2CC"),
                     (f"Строка только в {nm.a_in}", "F8CBAD"),
                     (f"Строка только в {nm.b_in}", "BDD7EE")]
        ws = writer.sheets["Описание различий"]
        style_table(ws, 1, wrap=("D",))
        color_rows(ws, "B", "D", row_rules)
        ws = writer.sheets["Различия по ячейкам"]
        style_table(ws, 1)
        color_rows(ws, "A", excel_letter(len(det.columns) - 1),
                   [(groups[g], group_colors[g]) for g in (1, 2, 3)])

        ws = writer.sheets["Столбцы"]
        style_table(ws, 1)
        ws.row_dimensions[1].height = 45
        titles = [c.value for c in ws[1]]
        for row in ws.iter_rows(min_row=2):
            for cell, title in zip(row, titles):
                if isinstance(cell.value, int):
                    cell.number_format = "#,##0"
                elif isinstance(cell.value, float):
                    cell.number_format = "#,##0.00####"
                cell.alignment = Alignment(horizontal="right") if title not in (
                    "Лист", "Столбец", "Замечания") else Alignment(wrap_text=True)
                group = next((g for g in (1, 2, 3) if groups[g] == title), None)
                if group and cell.value:
                    cell.fill = PatternFill("solid", fgColor=group_colors[group])
                elif title and "\n" in title and isinstance(cell.value, str):
                    cell.fill = PatternFill("solid", fgColor="FFEB9C")
                    cell.font = Font(bold=True)
        for cell in ws[1]:
            if cell.value in groups.values():
                ws.column_dimensions[cell.column_letter].width = 20
            elif cell.value == "Замечания":
                ws.column_dimensions[cell.column_letter].width = 60
        for name in ("Описание различий", "Различия по ячейкам", "Столбцы"):
            ws = writer.sheets[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions


# ---------- имена файлов и пакетный режим ----------

def is_timestamp(token):
    return token.isdigit() and len(token) >= 6


def name_tokens(path):
    return [t for t in Path(path).stem.split("_") if t and not is_timestamp(t)]


def derive_names(path1, path2, explicit=None):
    """Названия систем и общая часть имени из имён файлов."""
    t1, t2 = name_tokens(path1), name_tokens(path2)
    prefix, a, b = "", None, None
    if len(t1) == len(t2):
        diff = [k for k in range(len(t1)) if t1[k] != t2[k]]
        if len(diff) == 1:
            a, b = t1[diff[0]], t2[diff[0]]
            prefix = "_".join(t for k, t in enumerate(t1) if k != diff[0])
        elif not diff:
            prefix = "_".join(t1)
    if explicit:
        a, b = explicit
    if a and b:
        return Names(a, b, a, b, prefix)
    return Names(prefix=prefix)


def find_pairs(folder, systems):
    """Ищет в папке пары файлов, отличающихся в имени только названием системы."""
    files = [f for f in sorted(folder.glob("*.xlsx"))
             if not f.name.startswith("~$") and f.name != BATCH_SUMMARY
             and not f.stem.endswith("differences")]
    tokens = {f: name_tokens(f) for f in files}

    if not systems:
        votes = Counter()
        for i, f in enumerate(files):
            for g in files[i + 1:]:
                t1, t2 = tokens[f], tokens[g]
                if len(t1) == len(t2):
                    diff = [k for k in range(len(t1)) if t1[k] != t2[k]]
                    if len(diff) == 1:
                        votes[tuple(sorted((t1[diff[0]], t2[diff[0]])))] += 1
        if not votes:
            raise ValueError("В папке не найдено пар файлов. Имена должны отличаться "
                             "одним словом, например pks07_datacore_... и pks07_noogadev_...")
        top = votes.most_common(2)
        if len(top) > 1 and top[0][1] == top[1][1]:
            raise ValueError("Не удалось однозначно определить названия систем. "
                             "Укажите их: --systems datacore noogadev")
        systems = top[0][0]

    s1, s2 = systems
    groups = defaultdict(lambda: {s1: [], s2: []})
    unpaired = []
    for f in files:
        t = tokens[f]
        hits = [(k, s) for k, tok in enumerate(t) for s in (s1, s2) if tok == s]
        if len(hits) != 1:
            unpaired.append(f)
            continue
        k, s = hits[0]
        groups[tuple(t[:k] + ["*"] + t[k + 1:])][s].append(f)

    pairs, notes = [], []
    for key, g in sorted(groups.items()):
        if g[s1] and g[s2]:
            for s in (s1, s2):
                if len(g[s]) > 1:
                    notes.append(f"{'_'.join(key)}: несколько файлов {s}, взят последний "
                                 f"по имени ({max(g[s]).name})")
            pairs.append((max(g[s1]), max(g[s2])))
        else:
            unpaired += g[s1] + g[s2]
    return pairs, unpaired, notes


# ---------- запуск ----------

def print_result(res, file1, file2, report):
    print(f"\n{Path(file1).name}\n{Path(file2).name}")
    for c in res.checks:
        if c["Результат"] == INFO and not c["Комментарий"]:
            continue
        sheet = f"[{c['Лист']}] " if c["Лист"] else ""
        values = f": {c['a']} / {c['b']}" if c["a"] != "" or c["b"] != "" else ""
        comment = f" ({brief(c['Комментарий'], 110)})" if c["Комментарий"] else ""
        print(f"  {c['Результат']:<18}{sheet}{c['Проверка']}{values}{comment}")
    if FAST_READER is False:
        print("  Чтение файлов можно ускорить в несколько раз: "
              "python -m pip install python-calamine")
    print("Итог: " + ("ЕСТЬ РАЗЛИЧИЯ" if res.has_differences else "файлы совпадают"))
    print(f"Отчёт: {report}")


def run_pair(file1, file2, o, names, report):
    nm = derive_names(file1, file2, names)
    res = compare_files(file1, file2, o, nm)
    if report is None:   # по умолчанию отчёт кладём рядом с первым файлом
        report = Path(file1).parent / (f"{nm.prefix}_differences.xlsx" if nm.prefix
                                       else "differences.xlsx")
    elif report.is_dir():
        report = report / (f"{nm.prefix}_differences.xlsx" if nm.prefix
                           else f"{Path(file1).stem}_differences.xlsx")
    write_report(report, res, file1, file2, nm, o)
    print_result(res, file1, file2, report)
    return res, nm, report


def main():
    for stream in (sys.stdout, sys.stderr):   # чтобы консоль Windows не падала
        try:                                  # на символах вне её кодировки
            stream.reconfigure(errors="replace")
        except Exception:
            pass

    p = argparse.ArgumentParser(
        description="Сверка данных в двух Excel-файлах с одинаковой структурой. "
                    "Порядок строк и столбцов может быть разным.")
    p.add_argument("file1", type=Path, nargs="?", help="первый файл")
    p.add_argument("file2", type=Path, nargs="?", help="второй файл")
    p.add_argument("--dir", type=Path, metavar="ПАПКА",
                   help="сравнить все пары файлов в папке (вместо file1 и file2)")
    p.add_argument("--systems", nargs=2, metavar=("A", "B"),
                   help="названия систем в именах файлов для --dir, "
                        "например: datacore noogadev (по умолчанию определяются сами)")
    p.add_argument("--names", nargs=2, metavar=("A", "B"),
                   help="как называть файлы в отчёте (по умолчанию берётся из имён файлов)")
    p.add_argument("-o", "--output", type=Path,
                   help="файл отчёта или папка для отчётов "
                        "(по умолчанию <общая часть имени>_differences.xlsx)")
    p.add_argument("-k", "--key", nargs="+", metavar="СТОЛБЕЦ",
                   help="ключевые столбцы для сопоставления строк")
    p.add_argument("--by-order", action="store_true",
                   help="сравнивать строки строго по порядку")
    p.add_argument("-s", "--sheet", nargs="+", metavar="ЛИСТ",
                   help="сравнить только эти листы (по умолчанию все)")
    p.add_argument("--header", type=int, default=1, metavar="N",
                   help="номер строки с заголовками (по умолчанию 1)")
    p.add_argument("--skip-first-column", action="store_true",
                   help="не сравнивать столбец A (если в нём нумерация строк)")
    p.add_argument("--ignore", nargs="+", metavar="СТОЛБЕЦ", default=[],
                   help="столбцы, которые не нужно сравнивать")
    p.add_argument("--full", action="store_true",
                   help="добавить в отчёт подробные листы: все проверки, описание "
                        "различий, различия по ячейкам, столбцы")
    p.add_argument("--sum", nargs="+", metavar="СТОЛБЕЦ", default=[],
                   help="посчитать суммы по столбцам; указываются буквы столбцов Excel "
                        "первого файла, например: --sum AE AH или диапазон --sum AK:AM. "
                        "Без этого параметра ничего не суммируется")
    p.add_argument("-t", "--tolerance", type=float, default=0.0,
                   help="допустимая разница для чисел (по умолчанию 0)")
    p.add_argument("--ignore-case", action="store_true",
                   help="не учитывать регистр в тексте")
    p.add_argument("--text-numbers", action="store_true",
                   help="считать текст, похожий на число, числом: "
                        "'2025' = 2025, '0.00' = 0, '0810' = 810")
    p.add_argument("--similarity", type=float, default=0.5, metavar="ДОЛЯ",
                   help="какая доля ячеек должна совпасть, чтобы две строки считались "
                        "одной строкой с изменениями (от 0 до 1, по умолчанию 0.5)")
    args = p.parse_args()
    if not 0 <= args.similarity <= 1:
        p.error("--similarity должен быть от 0 до 1")
    if args.tolerance < 0:
        p.error("--tolerance не может быть отрицательным")
    if args.header < 1:
        p.error("--header: номер строки заголовков считается с 1")
    if args.key and args.by_order:
        p.error("укажите что-то одно: --key или --by-order")

    o = Options(keys=args.key, by_order=args.by_order, header=args.header - 1,
                tolerance=args.tolerance, ignore_case=args.ignore_case,
                text_numbers=args.text_numbers, similarity=args.similarity,
                ignore=tuple(c.strip() for c in args.ignore),
                skip_first_column=args.skip_first_column, sheets=args.sheet,
                full=args.full)

    try:
        o.sum_cols = parse_columns(args.sum)
        if args.dir:
            if args.file1 or args.file2:
                p.error("укажите либо два файла, либо --dir")
            if o.sum_cols:
                raise ValueError("--sum работает только при сравнении двух файлов: в папке "
                                 "у разных отчётов разные столбцы")
            if not args.dir.is_dir():
                raise ValueError(f"Папка не найдена: {args.dir}")
            out = args.output or args.dir
            if not out.is_dir():
                raise ValueError(f"Для --dir параметр -o должен быть папкой: {out}")
            pairs, unpaired, notes = find_pairs(args.dir, args.systems)
            summary, any_diff, any_error = [], False, False
            for f1, f2 in pairs:
                row = {"Файл 1": f1.name, "Файл 2": f2.name}
                try:
                    res, nm, report = run_pair(f1, f2, o, args.names, out)
                    t = res.totals
                    any_diff |= res.has_differences
                    row = {"Отчёт": nm.prefix, **row,
                           "Итог": "есть различия" if res.has_differences else "совпадают",
                           "Строк в файле 1": t["rows_a"], "Строк в файле 2": t["rows_b"],
                           "Изменено строк": t["changed"],
                           "Только в файле 1": t["only_a"], "Только в файле 2": t["only_b"],
                           "Столбцов отличается целиком": t["const_columns"],
                           "Отличий в структуре": t["structure"],
                           "Файл с подробностями": report.name}
                except Exception as e:
                    any_error = True
                    row["Итог"] = f"ошибка: {e}"
                    print(f"\n{f1.name}\n{f2.name}\n  Ошибка: {e}")
                summary.append(row)
            for f in unpaired:
                summary.append({"Файл 1": f.name, "Итог": "пара не найдена"})
            if summary:
                with pd.ExcelWriter(out / BATCH_SUMMARY, engine="openpyxl") as writer:
                    pd.DataFrame(summary).to_excel(writer, sheet_name="Итог", index=False)
                    style_table(writer.sheets["Итог"], 1)
            print(f"\nСравнено пар: {len(pairs)}. Общий итог: {out / BATCH_SUMMARY}")
            for n in notes:
                print("  Внимание: " + n)
            for f in unpaired:
                print(f"  Пара не найдена: {f.name}")
            if any_error:
                print("  Часть пар сравнить не удалось, см. ошибки выше", file=sys.stderr)
                return 2
            return 1 if any_diff else 0

        if not (args.file1 and args.file2):
            p.error("укажите два файла или папку через --dir")
        for f in (args.file1, args.file2):
            if not f.is_file():
                raise ValueError(f"Файл не найден: {f}")
        res, _, _ = run_pair(args.file1, args.file2, o, args.names, args.output)
        return 1 if res.has_differences else 0
    except PermissionError as e:
        print(f"Ошибка: нет доступа к файлу {e.filename}. Если он открыт в Excel, "
              "закройте его и запустите скрипт снова", file=sys.stderr)
        return 2
    except (ValueError, OSError) as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        return 2
    except Exception as e:   # непредвиденная ошибка: код 2, а не «различия есть»
        print(f"Ошибка ({type(e).__name__}): {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
