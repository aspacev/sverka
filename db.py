"""База данных сайта сверок: один файл SQLite."""

import sqlite3

SCHEMA = """
CREATE TABLE IF NOT EXISTS systems (
    id   INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS reports (
    id          INTEGER PRIMARY KEY,
    system_id   INTEGER NOT NULL REFERENCES systems(id),   -- у каждой системы свои отчёты
    name        TEXT NOT NULL,
    date_column TEXT,                          -- столбец с датой (для разбивки по месяцам)
    date_source TEXT NOT NULL DEFAULT 'auto',  -- auto: найден сайтом, manual: указан вручную
    UNIQUE (system_id, name)
);

CREATE TABLE IF NOT EXISTS runs (
    id          INTEGER PRIMARY KEY,
    created_at  TEXT NOT NULL,
    report_id   INTEGER NOT NULL REFERENCES reports(id),
    system_id   INTEGER NOT NULL REFERENCES systems(id),   -- система-источник данных
    period_from TEXT NOT NULL,                 -- ГГГГ-ММ-ДД; пусто, если период неизвестен
    period_to   TEXT NOT NULL,
    period_auto INTEGER DEFAULT 0,             -- 1: период взят из дат в файле
    status      TEXT NOT NULL,                 -- ok / diff / error
    error_text  TEXT,
    file_a      TEXT,                          -- имена загруженных файлов
    file_b      TEXT,
    options     TEXT,                          -- настройки сверки (JSON)
    rows_a      INTEGER DEFAULT 0,
    rows_b      INTEGER DEFAULT 0,
    exact       INTEGER DEFAULT 0,
    changed     INTEGER DEFAULT 0,
    only_a      INTEGER DEFAULT 0,
    only_b      INTEGER DEFAULT 0,
    cells       INTEGER DEFAULT 0,
    structure   INTEGER DEFAULT 0,             -- отличия в структуре и столбцы «целиком»
    by_month    INTEGER DEFAULT 0,             -- 1: расхождения разложены по месяцам
    date_column TEXT,
    unplaced    INTEGER DEFAULT 0,             -- строк с расхождениями без месяца
    summary     TEXT                           -- проверки, таблицы, структура (JSON)
);

CREATE TABLE IF NOT EXISTS run_months (
    run_id   INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    month    TEXT NOT NULL,                    -- ГГГГ-ММ
    bad_rows INTEGER NOT NULL DEFAULT 0,
    partial  INTEGER NOT NULL DEFAULT 0,       -- 1: период покрывает месяц не целиком
    PRIMARY KEY (run_id, month)
);

CREATE TABLE IF NOT EXISTS run_diffs (
    id     INTEGER PRIMARY KEY,
    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    sheet  TEXT NOT NULL,
    kind   TEXT NOT NULL,                      -- changed / only_a / only_b
    row_a  INTEGER,
    row_b  INTEGER,
    month  TEXT,
    vals_a TEXT,                               -- значения строки (JSON)
    vals_b TEXT,
    kinds  TEXT,                               -- номер столбца -> вид отличия (JSON)
    hint   TEXT
);

CREATE INDEX IF NOT EXISTS runs_report ON runs(report_id, created_at);
CREATE INDEX IF NOT EXISTS diffs_run ON run_diffs(run_id, month);
"""


def connect(path):
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init(path):
    conn = connect(path)
    # Ранние версии базы: две системы в сверке, отчёты без системы.
    # Пустые таблицы пересоздаём; базу с данными не трогаем
    runs = [row[1] for row in conn.execute("PRAGMA table_info(runs)")]
    reports = [row[1] for row in conn.execute("PRAGMA table_info(reports)")]
    if "system_a_id" in runs or (reports and "system_id" not in reports):
        if any(conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
               for t in ("runs", "reports")):
            raise RuntimeError("База создана старой версией сайта и содержит данные. "
                               "Переименуйте файл базы и запустите сайт снова.")
        conn.executescript("DROP TABLE run_diffs; DROP TABLE run_months; DROP TABLE runs; "
                           "DROP TABLE reports;")
    conn.executescript(SCHEMA)
    # Столбцы, появившиеся в новых версиях: добавляем в существующую базу без потери данных
    if "period_auto" not in [row[1] for row in conn.execute("PRAGMA table_info(runs)")]:
        conn.execute("ALTER TABLE runs ADD COLUMN period_auto INTEGER DEFAULT 0")
    conn.commit()
    conn.close()
