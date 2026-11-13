"""
Сайт для сверки Excel-отчётов.

Запуск:  docker compose up -d   или   python app.py   (сайт: http://localhost:5000)
Данные:  папка data рядом с этим файлом (база sverka.db, загруженные файлы, отчёты);
         другая папка задаётся переменной SVERKA_DATA
"""

import json
import os
import re
import shutil
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

from flask import (Flask, abort, flash, g, redirect, render_template, request, send_file,
                   url_for)

import compare_excel as ce
import db as dbm
import service

PAGE = 50
# Два хранилища данных (DWH), между которыми идёт сверка. Первое - эталон
DWH_A, DWH_B = "datacore", "nooga"
RUNS = """
    SELECT r.*, rp.name AS report, s.name AS system
    FROM runs r
    JOIN reports rp ON rp.id = r.report_id
    JOIN systems s ON s.id = r.system_id
"""


def parse_date(text):
    try:
        return date.fromisoformat((text or "").strip())
    except ValueError:
        return None


def plural_diff(n):
    return f"{n} {ce.plural(n, 'расхождение', 'расхождения', 'расхождений')}"


def safe_name(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", Path(name or "").name).strip(" .")
    return name or "file.xlsx"


def lines(text):
    return list(dict.fromkeys(x.strip() for x in (text or "").splitlines() if x.strip()))


def create_app(data_dir=None):
    app = Flask(__name__)
    app.secret_key = os.urandom(16)
    data = Path(data_dir or os.environ.get("SVERKA_DATA") or Path(__file__).parent / "data")
    data.mkdir(parents=True, exist_ok=True)
    dbm.init(data / "sverka.db")

    def db():
        if "db" not in g:
            g.db = dbm.connect(data / "sverka.db")
        return g.db

    @app.teardown_appcontext
    def close_db(_):
        conn = g.pop("db", None)
        if conn is not None:
            conn.close()

    @app.template_filter("d")
    def format_date(text):
        d = parse_date(text)
        return d.strftime("%d.%m.%Y") if d else text

    @app.template_filter("dt")
    def format_datetime(text):
        try:
            return datetime.fromisoformat(text).strftime("%d.%m.%Y %H:%M")
        except (TypeError, ValueError):
            return text

    @app.template_filter("num")
    def format_number(n):
        return f"{n or 0:,}".replace(",", " ")

    app.jinja_env.globals.update(plural_diff=plural_diff, month_title=service.month_title,
                                 dwh_a=DWH_A, dwh_b=DWH_B)

    def lists():
        return (db().execute("SELECT * FROM systems ORDER BY name").fetchall(),
                db().execute("SELECT p.*, s.name AS system FROM reports p "
                             "JOIN systems s ON s.id = p.system_id "
                             "ORDER BY s.name, p.name").fetchall())

    def run_label(run):
        """Короткий итог сверки: (вид, текст)."""
        if run["status"] == "error":
            return "err", "ошибка"
        if run["status"] == "ok":
            return "ok", "совпало"
        rows = run["changed"] + run["only_a"] + run["only_b"]
        return "bad", plural_diff(rows) if rows else "расхождения"

    app.jinja_env.globals.update(run_label=run_label)

    # ---------- новая сверка ----------

    @app.route("/")
    def index():
        return redirect(url_for("grid"))

    @app.route("/new", methods=["GET", "POST"])
    def new():
        systems, reports = lists()
        form = request.form if request.method == "POST" else request.args
        errors = []
        if request.method == "POST":
            report = db().execute("SELECT * FROM reports WHERE id = ?",
                                  (form.get("report"),)).fetchone()
            system = db().execute("SELECT * FROM systems WHERE id = ?",
                                  (form.get("system"),)).fetchone()
            d_from, d_to = parse_date(form.get("from")), parse_date(form.get("to"))
            file_a, file_b = request.files.get("file_a"), request.files.get("file_b")
            if not report:
                errors.append("Выберите отчёт.")
            if not system:
                errors.append("Выберите систему.")
            elif report and report["system_id"] != system["id"]:
                errors.append("Этот отчёт относится к другой системе. Выберите отчёт заново.")
            # Период необязателен: без него он берётся из дат в самих файлах
            has_from = bool((form.get("from") or "").strip())
            has_to = bool((form.get("to") or "").strip())
            if has_from != has_to or (has_from and not (d_from and d_to)):
                errors.append("Укажите обе даты периода или оставьте обе пустыми.")
            elif d_from and d_from > d_to:
                errors.append("Дата «с» позже даты «по».")
            for f, title in ((file_a, DWH_A), (file_b, DWH_B)):
                if not f or not f.filename:
                    errors.append(f"Выберите файл из {title}.")
                elif not f.filename.lower().endswith((".xlsx", ".xlsm")):
                    errors.append(f"Файл из {title} должен быть Excel (.xlsx).")
            try:
                header = int(form.get("header") or 1)
                if header < 1:
                    raise ValueError
            except ValueError:
                header = 1
                errors.append("Номер строки заголовков — целое число от 1.")

            if not errors:
                options = {
                    "keys": [k.strip() for k in form.get("keys", "").split(";") if k.strip()]
                            or None,
                    "ignore": tuple(k.strip() for k in form.get("ignore", "").split(";")
                                    if k.strip()),
                    "skip_first_column": bool(form.get("skip_first")),
                    "text_numbers": bool(form.get("text_numbers")),
                    "header": header - 1,
                }
                cur = db().execute(
                    "INSERT INTO runs (created_at, report_id, system_id, "
                    "period_from, period_to, status, error_text, file_a, file_b, options) "
                    "VALUES (?, ?, ?, ?, ?, 'error', 'Сверка не завершена', ?, ?, ?)",
                    (datetime.now().isoformat(" ", "seconds"), report["id"], system["id"],
                     d_from.isoformat() if d_from else "", d_to.isoformat() if d_to else "",
                     safe_name(file_a.filename), safe_name(file_b.filename),
                     json.dumps(options, ensure_ascii=False)))
                run_id = cur.lastrowid
                db().commit()
                paths = []
                for side, f in (("a", file_a), ("b", file_b)):
                    folder = data / "uploads" / str(run_id) / side
                    folder.mkdir(parents=True, exist_ok=True)
                    paths.append(folder / safe_name(f.filename))
                    f.save(paths[-1])
                service.execute(db(), data, run_id, paths[0], paths[1], DWH_A, DWH_B,
                                report, d_from, d_to, options,
                                title=f"{report['name']}, {system['name']}")
                return redirect(url_for("run", run_id=run_id))
        return render_template("new.html", systems=systems, reports=reports, form=form,
                               errors=errors, active="new")

    # ---------- одна сверка ----------

    def get_run(run_id):
        run = db().execute(RUNS + " WHERE r.id = ?", (run_id,)).fetchone()
        if run is None:
            abort(404)
        return run

    @app.route("/runs/<int:run_id>")
    def run(run_id):
        run = get_run(run_id)
        summary = json.loads(run["summary"] or "{}")
        months = db().execute("SELECT * FROM run_months WHERE run_id = ? ORDER BY month",
                              (run_id,)).fetchall()
        month = request.args.get("month", "")
        where, params = "run_id = ?", [run_id]
        if month == "none":
            where += " AND month IS NULL"
        elif month:
            where += " AND month = ?"
            params.append(month)
        total = db().execute(f"SELECT COUNT(*) FROM run_diffs WHERE {where}", params).fetchone()[0]
        page = max(1, request.args.get("page", 1, type=int))
        rows = db().execute(f"SELECT * FROM run_diffs WHERE {where} ORDER BY id LIMIT ? OFFSET ?",
                            (*params, PAGE, (page - 1) * PAGE)).fetchall()
        tables = {t["sheet"]: t for t in summary.get("tables", [])}
        groups = []
        for row in rows:
            if not groups or groups[-1]["sheet"] != row["sheet"]:
                t = tables.get(row["sheet"], {"columns": [], "tech": 0, "whole": 0})
                n = len(t["columns"])
                groups.append({"sheet": row["sheet"], "columns": t["columns"],
                               "first_tech": n - t["tech"],
                               "first_whole": n - t["tech"] - t["whole"], "rows": []})
            groups[-1]["rows"].append({
                "kind": row["kind"], "row_a": row["row_a"], "row_b": row["row_b"],
                "month": row["month"], "hint": row["hint"],
                "a": json.loads(row["vals_a"]) if row["vals_a"] else None,
                "b": json.loads(row["vals_b"]) if row["vals_b"] else None,
                "kinds": json.loads(row["kinds"] or "{}")})
        stored = db().execute("SELECT COUNT(*) FROM run_diffs WHERE run_id = ?",
                              (run_id,)).fetchone()[0]
        return render_template(
            "run.html", run=run, summary=summary, months=months, month=month, groups=groups,
            total=total, page=page, pages=max(1, -(-total // PAGE)), stored=stored,
            bad_rows=run["changed"] + run["only_a"] + run["only_b"],
            many_sheets=len(tables) > 1, active="history")

    @app.route("/runs/<int:run_id>/excel")
    def run_excel(run_id):
        run = get_run(run_id)
        path = data / "reports" / f"run_{run_id}.xlsx"
        if not path.is_file():
            abort(404)
        period = (f"{run['period_from']}_{run['period_to']}" if run["period_from"]
                  else run["created_at"][:10])
        name = safe_name(f"{run['report']}_{run['system']}_{period}.xlsx")
        return send_file(path, as_attachment=True, download_name=name)

    @app.route("/runs/<int:run_id>/delete", methods=["POST"])
    def run_delete(run_id):
        get_run(run_id)
        db().execute("DELETE FROM runs WHERE id = ?", (run_id,))
        db().commit()
        shutil.rmtree(data / "uploads" / str(run_id), ignore_errors=True)
        (data / "reports" / f"run_{run_id}.xlsx").unlink(missing_ok=True)
        flash("Сверка удалена.")
        return redirect(url_for("history"))

    # ---------- история ----------

    @app.route("/history")
    def history():
        systems, reports = lists()
        where, params = [], []
        if request.args.get("report"):
            where.append("r.report_id = ?")
            params.append(request.args["report"])
        if request.args.get("system"):
            where.append("r.system_id = ?")
            params.append(request.args["system"])
        if request.args.get("status") in ("ok", "diff", "error"):
            where.append("r.status = ?")
            params.append(request.args["status"])
        clause = (" WHERE " + " AND ".join(where)) if where else ""
        total = db().execute("SELECT COUNT(*) FROM runs r" + clause, params).fetchone()[0]
        page = max(1, request.args.get("page", 1, type=int))
        runs = db().execute(RUNS + clause + " ORDER BY r.id DESC LIMIT ? OFFSET ?",
                            (*params, PAGE, (page - 1) * PAGE)).fetchall()
        args = {k: v for k, v in request.args.items() if k != "page" and v}
        return render_template("history.html", runs=runs, systems=systems, reports=reports,
                               total=total, page=page, pages=max(1, -(-total // PAGE)),
                               args=args, active="history")

    # ---------- сетка ----------

    @app.route("/grid")
    def grid():
        systems, reports = lists()
        system = request.args.get("system", type=int)
        if system not in {s["id"] for s in systems}:
            last = db().execute("SELECT system_id FROM runs ORDER BY id DESC LIMIT 1").fetchone()
            system = last[0] if last else systems[0]["id"] if systems else None
        count = request.args.get("n", 6, type=int)
        count = count if count in (6, 12) else 6
        end = request.args.get("end", "")
        if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", end):
            end = service.month_key(date.today())
        months = [service.shift_month(end, k - count + 1) for k in range(count)]
        first_day = service.month_bounds(months[0])[0].isoformat()
        last_day = service.month_bounds(months[-1])[1].isoformat()

        # Для каждой клетки (отчёт, месяц) берём самую свежую сверку этой системы
        cells = {}
        runs = db().execute(
            "SELECT * FROM runs WHERE status != 'error' AND period_from <= ? AND period_to >= ? "
            "AND system_id = ? ORDER BY id", (last_day, first_day, system)).fetchall()
        for r in runs:
            total = r["changed"] + r["only_a"] + r["only_b"]
            other = r["status"] == "diff" and (r["unplaced"] or r["structure"] or not total)
            for m in db().execute("SELECT * FROM run_months WHERE run_id = ?", (r["id"],)):
                if m["month"] not in months:
                    continue
                bad = m["bad_rows"] > 0 or other
                if bad and m["bad_rows"]:
                    text = str(m["bad_rows"])
                    tip = plural_diff(m["bad_rows"]) + " за месяц"
                    if other:
                        text += "+"
                        tip += "; есть расхождения без даты или в структуре"
                elif bad:
                    text, tip = "!", "Есть расхождения без даты или в структуре файла"
                elif m["partial"]:
                    text, tip = "часть", "Сверена часть месяца, расхождений нет"
                else:
                    text, tip = "совпало", "Расхождений нет"
                if m["partial"] and bad:
                    tip += "; сверена часть месяца"
                cells[(r["report_id"], m["month"])] = {
                    "run": r, "kind": "bad" if bad else "part" if m["partial"] else "ok",
                    "text": text, "tip": tip,
                    "wide": r["status"] == "diff" and not r["by_month"]}

        rows, done, clean = [], 0, 0
        for rp in (r for r in reports if r["system_id"] == system):
            line, states, k = [], set(), 0
            while k < len(months):
                cell = cells.get((rp["id"], months[k]))
                if cell is None:
                    first, last = service.month_bounds(months[k])
                    line.append({"kind": "no", "span": 1, "from": first.isoformat(),
                                 "to": last.isoformat()})
                    states.add("none")
                    k += 1
                    continue
                span = 1
                if cell["wide"]:        # одна сверка без разбивки: общая полоса
                    while (k + span < len(months)
                           and (other := cells.get((rp["id"], months[k + span])))
                           and other["run"]["id"] == cell["run"]["id"]):
                        span += 1
                    r = cell["run"]
                    total = r["changed"] + r["only_a"] + r["only_b"]
                    cell = dict(cell, text=(plural_diff(total) if total else "расхождения")
                                + " · без разбивки по месяцам",
                                tip="В отчёте не найден столбец с датой")
                states.add(cell["kind"])
                line.append(dict(cell, span=span))
                k += span
            done += "none" not in states
            clean += states == {"ok"}
            rows.append({"report": rp, "cells": line, "states": " ".join(sorted(states))})
        return render_template(
            "grid.html", systems=systems, reports=reports, rows=rows, months=months,
            system=system, n=count, end=end, done=done, clean=clean,
            prev=service.shift_month(end, -count), next=service.shift_month(end, count),
            today=service.month_key(date.today()), active="grid")

    # ---------- справочники ----------

    @app.route("/dict")
    def dictionaries():
        systems = db().execute(
            "SELECT s.*, (SELECT COUNT(*) FROM reports p WHERE p.system_id = s.id) AS used "
            "FROM systems s ORDER BY name").fetchall()
        reports = db().execute(
            "SELECT p.*, (SELECT COUNT(*) FROM runs r WHERE r.report_id = p.id) AS used "
            "FROM reports p ORDER BY name").fetchall()
        chosen = request.args.get("system", type=int)
        if chosen not in {s["id"] for s in systems}:
            chosen = systems[0]["id"] if systems else None
        return render_template("dict.html", systems=systems, active="dict", chosen=chosen,
                               reports=[r for r in reports if r["system_id"] == chosen])

    @app.route("/dict/<kind>/add", methods=["POST"])
    def dict_add(kind):
        if kind not in ("systems", "reports"):
            abort(404)
        system = None
        if kind == "reports":
            system = db().execute("SELECT * FROM systems WHERE id = ?",
                                  (request.form.get("system"),)).fetchone()
            if system is None:
                flash("Выберите систему, к которой относятся отчёты.")
                return redirect(url_for("dictionaries") + "#reports")
        added = 0
        for name in lines(request.form.get("names")):
            try:
                if system is None:
                    db().execute("INSERT INTO systems (name) VALUES (?)", (name,))
                else:
                    db().execute("INSERT INTO reports (system_id, name) VALUES (?, ?)",
                                 (system["id"], name))
                added += 1
            except sqlite3.IntegrityError:
                flash(f"«{name}» уже есть в списке.")
        db().commit()
        if added:
            flash(f"Добавлено: {added}.")
        if system is not None:
            return redirect(url_for("dictionaries", system=system["id"]) + "#reports")
        return redirect(url_for("dictionaries") + "#systems")

    @app.route("/dict/<kind>/<int:item_id>/save", methods=["POST"])
    def dict_save(kind, item_id):
        if kind not in ("systems", "reports"):
            abort(404)
        name = request.form.get("name", "").strip()
        back = url_for("dictionaries", system=request.form.get("system") or None) + "#" + kind
        if not name:
            flash("Название не может быть пустым.")
            return redirect(back)
        try:
            if kind == "systems":
                db().execute("UPDATE systems SET name = ? WHERE id = ?", (name, item_id))
            else:
                column = request.form.get("date_column", "").strip()
                old = db().execute("SELECT * FROM reports WHERE id = ?", (item_id,)).fetchone()
                if old is None:
                    abort(404)
                if column == (old["date_column"] or ""):
                    source = old["date_source"]
                else:
                    source = "manual" if column else "auto"
                db().execute("UPDATE reports SET name = ?, date_column = ?, date_source = ? "
                             "WHERE id = ?", (name, column or None, source, item_id))
            db().commit()
            flash("Сохранено.")
        except sqlite3.IntegrityError:
            flash(f"«{name}» уже есть в списке.")
        return redirect(back)

    @app.route("/dict/<kind>/<int:item_id>/delete", methods=["POST"])
    def dict_delete(kind, item_id):
        if kind not in ("systems", "reports"):
            abort(404)
        back = url_for("dictionaries", system=request.form.get("system") or None) + "#" + kind
        try:
            db().execute(f"DELETE FROM {kind} WHERE id = ?", (item_id,))
            db().commit()
            flash("Удалено.")
        except sqlite3.IntegrityError:
            db().rollback()
            flash("Нельзя удалить систему, пока у неё есть отчёты." if kind == "systems" else
                  "Нельзя удалить отчёт, пока по нему есть сверки. Сначала удалите их в истории.")
        return redirect(back)

    return app


def main():
    """Запуск сайта. Адрес и порт задаются переменными HOST и PORT."""
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", 5000))
    app = create_app()
    try:
        from waitress import serve
    except ImportError:            # waitress не установлен: встроенный сервер Flask
        app.run(host=host, port=port)
    else:
        shown = "localhost" if host == "0.0.0.0" else host
        print(f"Сайт сверок запущен: http://{shown}:{port}", flush=True)
        serve(app, host=host, port=port, threads=8)


if __name__ == "__main__":
    main()
