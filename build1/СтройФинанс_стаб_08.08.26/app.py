from __future__ import annotations

import argparse
import json
import mimetypes
import os
import shutil
import sqlite3
import sys
import threading
import time
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

from xlsx_export import build_page_workbook, build_workbook
from xlsx_import import parse_workbook
from docx_export import build_exec_docx

APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"
DATA_DIR = APP_DIR / "data"
DB_PATH = DATA_DIR / "business.db"
SEED_PATH = APP_DIR / "seed_data.json"
HOST = "127.0.0.1"
DOCUMENTS_DIR = DATA_DIR / "documents"


def documents_dir(project_id: int) -> Path:
    return DOCUMENTS_DIR / str(project_id)


def sanitize_filename(name: str) -> str:
    name = (name or "file").replace("\\", "_").replace("/", "_")
    cleaned = "".join(ch if ch.isalnum() or ch in "._- " else "_" for ch in name)
    cleaned = " ".join(cleaned.split())
    return cleaned[:120] or "file"


def document_path(row) -> Path:
    return documents_dir(int(row["project_id"])) / f"{int(row['id'])}_{row['filename']}"


def remove_document_files(row) -> None:
    try:
        document_path(row).unlink(missing_ok=True)
    except OSError:
        pass


def remove_act_documents(conn, act_ids) -> None:
    if not act_ids:
        return
    rows = conn.execute(
        f"SELECT * FROM documents WHERE act_id IN ({','.join('?' * len(act_ids))})", act_ids
    ).fetchall()
    for row in rows:
        remove_document_files(row)
        conn.execute("DELETE FROM documents WHERE id=?", (row["id"],))

PROJECT_SECTIONS = {"revenue", "cashless", "asset", "cash", "payroll"}
OVERHEAD_GROUPS = {"cashless", "cash"}
RESERVED_NAMES = {"накладні", "результат"}


def connect() -> sqlite3.Connection:
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    first_run = not DB_PATH.exists()
    with connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                sort_order INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS project_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
                section TEXT NOT NULL,
                category TEXT NOT NULL DEFAULT '',
                description TEXT NOT NULL,
                amount REAL NOT NULL DEFAULT 0,
                date TEXT NOT NULL DEFAULT '',
                note TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS overhead_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_type TEXT NOT NULL,
                category TEXT NOT NULL DEFAULT '',
                description TEXT NOT NULL,
                amount REAL NOT NULL DEFAULT 0,
                date TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS contracts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
                name TEXT NOT NULL,
                date TEXT NOT NULL DEFAULT '',
                amount REAL NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS contract_acts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                contract_id INTEGER NOT NULL REFERENCES contracts(id) ON DELETE CASCADE,
                number TEXT NOT NULL DEFAULT '',
                date TEXT NOT NULL DEFAULT '',
                amount REAL NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
                act_id INTEGER REFERENCES contract_acts(id) ON DELETE CASCADE,
                filename TEXT NOT NULL,
                size INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            """
        )
        cols = [r[1] for r in conn.execute("PRAGMA table_info(documents)")]
        if "act_id" not in cols:
            conn.execute("ALTER TABLE documents ADD COLUMN act_id INTEGER REFERENCES contract_acts(id) ON DELETE CASCADE")
        conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('conversion_rate', '0.14')")
        count = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
        if (first_run or count == 0) and SEED_PATH.exists():
            seed = json.loads(SEED_PATH.read_text(encoding="utf-8"))
            restore_state(conn, seed, reset=True)


def clean_name(value: object) -> str:
    name = " ".join(str(value or "").strip().split())
    if not name:
        raise ValueError("Введите название стройки")
    if name.casefold() in RESERVED_NAMES:
        raise ValueError("Название зарезервировано для служебного листа")
    if len(name) > 120:
        raise ValueError("Название слишком длинное")
    return name


def clean_amount(value: object) -> float:
    if isinstance(value, str):
        value = value.replace(" ", "").replace(",", ".")
    try:
        amount = float(value)
    except (TypeError, ValueError):
        raise ValueError("Сумма должна быть числом")
    if amount < 0:
        raise ValueError("Сумма не может быть отрицательной")
    if amount > 1_000_000_000_000:
        raise ValueError("Сумма слишком большая")
    return round(amount, 2)


def clean_date(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        raise ValueError("Дата должна быть в формате ГГГГ-ММ-ДД")


def clean_text(value: object, error: str) -> str:
    text = " ".join(str(value or "").strip().split())
    if not text:
        raise ValueError(error)
    if len(text) > 250:
        raise ValueError("Название слишком длинное")
    return text


def restore_state(conn: sqlite3.Connection, data: dict, reset: bool = True) -> None:
    if reset:
        conn.execute("DELETE FROM project_entries")
        conn.execute("DELETE FROM overhead_entries")
        conn.execute("DELETE FROM projects")
        conn.execute("DELETE FROM documents")
    settings = data.get("settings", {})
    rate = float(settings.get("conversion_rate", 0.14) or 0.14)
    if not 0 <= rate <= 1:
        rate = 0.14
    conn.execute(
        "INSERT INTO settings(key, value) VALUES('conversion_rate', ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (str(rate),),
    )
    project_id_map: dict[object, int] = {}
    for order, project in enumerate(data.get("projects", []), start=1):
        name = clean_name(project.get("name"))
        cur = conn.execute("INSERT INTO projects(name, sort_order) VALUES(?, ?)", (name, order))
        new_id = int(cur.lastrowid)
        project_id_map[project.get("id", new_id)] = new_id
        for entry in project.get("entries", []):
            section = str(entry.get("section", "")).strip()
            if section not in PROJECT_SECTIONS:
                continue
            description = str(entry.get("description", "")).strip()
            amount = clean_amount(entry.get("amount", 0))
            conn.execute(
                """INSERT INTO project_entries(project_id, section, category, description, amount, date, note)
                   VALUES(?, ?, ?, ?, ?, ?, ?)""",
                (
                    new_id,
                    section,
                    str(entry.get("category", "")).strip(),
                    description,
                    amount,
                    clean_date(entry.get("date", "")),
                    str(entry.get("note", "")).strip(),
                ),
            )
    for entry in data.get("overheads", []):
        group = str(entry.get("group_type", "cashless"))
        if group not in OVERHEAD_GROUPS:
            group = "cashless"
        conn.execute(
            """INSERT INTO overhead_entries(group_type, category, description, amount, date)
               VALUES(?, ?, ?, ?, ?)""",
            (
                group,
                str(entry.get("category", "")).strip(),
                str(entry.get("description", "")).strip(),
                clean_amount(entry.get("amount", 0)),
                clean_date(entry.get("date", "")),
            ),
        )
    for contract in data.get("contracts", []):
        project_id = project_id_map.get(contract.get("project_id"))
        if project_id is None:
            continue
        cur = conn.execute(
            """INSERT INTO contracts(project_id, name, date, amount)
               VALUES(?, ?, ?, ?)""",
            (
                project_id,
                clean_text(contract.get("name"), "Название договора не может быть пустым"),
                clean_date(contract.get("date", "")),
                clean_amount(contract.get("amount", 0)),
            ),
        )
        contract_id = int(cur.lastrowid)
        for act in contract.get("acts", []):
            conn.execute(
                """INSERT INTO contract_acts(contract_id, number, date, amount)
                   VALUES(?, ?, ?, ?)""",
                (
                    contract_id,
                    clean_text(act.get("number"), "Номер акта не может быть пустым"),
                    clean_date(act.get("date", "")),
                    clean_amount(act.get("amount", 0)),
                ),
            )


def get_state(conn: sqlite3.Connection) -> dict:
    rate_row = conn.execute("SELECT value FROM settings WHERE key='conversion_rate'").fetchone()
    rate = float(rate_row[0]) if rate_row else 0.14
    projects = []
    for p in conn.execute("SELECT id, name FROM projects ORDER BY sort_order, id"):
        entries = [dict(row) for row in conn.execute(
            """SELECT id, section, category, description, amount, date, note
               FROM project_entries WHERE project_id=? ORDER BY date, id""",
            (p["id"],),
        )]
        projects.append({"id": p["id"], "name": p["name"], "entries": entries})
    overheads = [dict(row) for row in conn.execute(
        "SELECT id, group_type, category, description, amount, date FROM overhead_entries ORDER BY date, id"
    )]
    contracts = []
    for contract in conn.execute("SELECT id, project_id, name, date, amount FROM contracts ORDER BY date, id"):
        acts = [dict(row) for row in conn.execute(
            "SELECT id, number, date, amount FROM contract_acts WHERE contract_id=? ORDER BY date, id",
            (contract["id"],),
        )]
        contracts.append({**contract, "acts": acts})
    documents = [dict(row) for row in conn.execute(
        "SELECT id, project_id, act_id, filename, size, created_at AS date FROM documents ORDER BY created_at, id"
    )]
    state = {
        "settings": {"conversion_rate": rate},
        "projects": projects,
        "overheads": overheads,
        "contracts": contracts,
        "documents": documents,
    }
    state["results"] = calculate_results(state)
    return state


def calculate_results(state: dict) -> dict:
    rate = float(state.get("settings", {}).get("conversion_rate", 0.14) or 0)
    overheads = state.get("overheads", [])
    overhead_cashless = sum(float(e.get("amount", 0) or 0) for e in overheads if e.get("group_type") == "cashless")
    overhead_cash = sum(float(e.get("amount", 0) or 0) for e in overheads if e.get("group_type") == "cash")
    overhead_adjustment = overhead_cashless * rate
    distributable = overhead_cashless + overhead_cash - overhead_adjustment

    base_rows = []
    total_turnover = 0.0
    for project in state.get("projects", []):
        entries = project.get("entries", [])
        sums = {key: 0.0 for key in PROJECT_SECTIONS}
        for entry in entries:
            section = entry.get("section")
            if section in sums:
                sums[section] += float(entry.get("amount", 0) or 0)
        turnover = sums["revenue"]
        cashless = sums["cashless"] + sums["asset"]
        balance = turnover - cashless
        conversion = balance * rate
        direct_cash = sums["cash"] + sums["payroll"]
        profit_before_overhead = turnover - cashless - conversion - direct_cash
        base_rows.append({
            "project_id": project["id"],
            "name": project["name"],
            "turnover": turnover,
            "cashless": cashless,
            "balance": balance,
            "conversion": conversion,
            "cash_expenses": direct_cash,
            "profit_before_overhead": profit_before_overhead,
        })
        total_turnover += turnover

    rows = []
    for row in base_rows:
        allocated = distributable * row["turnover"] / total_turnover if total_turnover else 0.0
        net = row["profit_before_overhead"] - allocated
        percent = net / row["turnover"] if row["turnover"] else 0.0
        rows.append({**row, "allocated_overhead": allocated, "net_profit": net, "profit_percent": percent})

    totals = {
        "turnover": sum(r["turnover"] for r in rows),
        "allocated_overhead": distributable,
        "cashless": sum(r["cashless"] for r in rows),
        "conversion": sum(r["conversion"] for r in rows),
        "cash_expenses": sum(r["cash_expenses"] for r in rows),
        "net_profit": sum(r["net_profit"] for r in rows),
        "profit_percent": 0.0,
        "overhead_cashless": overhead_cashless,
        "overhead_cash": overhead_cash,
        "overhead_adjustment": overhead_adjustment,
        "overhead_coefficient": distributable / total_turnover if total_turnover else 0.0,
    }
    totals["profit_percent"] = totals["net_profit"] / totals["turnover"] if totals["turnover"] else 0.0
    return {"rows": rows, "totals": totals}


class Handler(BaseHTTPRequestHandler):
    server_version = "StroyFinance/1.1"

    def log_message(self, fmt: str, *args) -> None:
        sys.stdout.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    def send_json(self, data: object, status: int = 200) -> None:
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(payload)

    def send_error_json(self, message: str, status: int = 400) -> None:
        self.send_json({"error": message}, status)

    def send_xlsx(self, payload: bytes, filename: str) -> None:
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(filename)}")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0") or 0)
        if length > 10_000_000:
            raise ValueError("Слишком большой запрос")
        raw = self.rfile.read(length) if length else b"{}"
        try:
            data = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ValueError("Некорректный JSON")
        if not isinstance(data, dict):
            raise ValueError("Ожидается объект JSON")
        return data

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        try:
            if path == "/api/state":
                with connect() as conn:
                    self.send_json(get_state(conn))
                return
            if path.startswith("/api/documents/") and path.endswith("/download"):
                doc_id = int(path.split("/")[3])
                with connect() as conn:
                    row = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
                if row is None or not document_path(row).exists():
                    self.send_error_json("Файл не найден", 404)
                    return
                payload = document_path(row).read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(row['filename'])}")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if path == "/api/export.xlsx":
                with connect() as conn:
                    state = get_state(conn)
                payload = build_workbook(state)
                filename = f"СтройФинанс_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
                self.send_xlsx(payload, filename)
                return
            if path in {"/api/export/dashboard.xlsx", "/api/export/overheads.xlsx", "/api/export/results.xlsx"}:
                page = path.split("/")[-1].removesuffix(".xlsx")
                with connect() as conn:
                    state = get_state(conn)
                payload = build_page_workbook(state, page)
                names = {"dashboard": "Обзор", "overheads": "Накладні", "results": "Результат"}
                filename = f"{names[page]}_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
                self.send_xlsx(payload, filename)
                return
            if path == "/api/export/projects.xlsx":
                query = parse_qs(parsed.query)
                project_id = int((query.get("project_id") or ["0"])[0] or 0)
                with connect() as conn:
                    state = get_state(conn)
                payload = build_page_workbook(state, "project", project_id)
                project = next((p for p in state.get("projects", []) if p.get("id") == project_id), None)
                name = project["name"] if project else "Стройка"
                filename = f"{name}_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
                self.send_xlsx(payload, filename)
                return
            if path == "/api/export/contracts.xlsx":
                query = parse_qs(parsed.query)
                project_id = int((query.get("project_id") or ["0"])[0] or 0)
                with connect() as conn:
                    state = get_state(conn)
                payload = build_page_workbook(state, "contracts", project_id)
                project = next((p for p in state.get("projects", []) if p.get("id") == project_id), None)
                name = project["name"] if project else "Стройка"
                filename = f"{name}_Договори_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
                self.send_xlsx(payload, filename)
                return
            if path == "/api/export/execdoc.docx":
                query = parse_qs(parsed.query)
                project_id = int((query.get("project_id") or ["0"])[0] or 0)
                with connect() as conn:
                    state = get_state(conn)
                payload = build_exec_docx(state, project_id)
                project = next((p for p in state.get("projects", []) if p.get("id") == project_id), None)
                name = project["name"] if project else "Стройка"
                filename = f"{name}_Виконавча_документація_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.docx"
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(filename)}")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if path == "/api/backup.json":
                with connect() as conn:
                    state = get_state(conn)
                state.pop("results", None)
                payload = json.dumps(state, ensure_ascii=False, indent=2).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Disposition", "attachment; filename=stroyfinance_backup.json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            self.serve_static(path)
        except Exception as exc:
            self.send_error_json(str(exc), 500)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        try:
            if path.startswith("/api/projects/") and path.endswith("/documents"):
                project_id = int(path.split("/")[3])
                query = parse_qs(parsed.query)
                filename = sanitize_filename((query.get("filename") or [""])[0])
                act_id_raw = (query.get("act_id") or [""])[0]
                act_id = int(act_id_raw) if act_id_raw.strip().isdigit() else None
                length = int(self.headers.get("Content-Length", "0") or 0)
                if length <= 0:
                    raise ValueError("Выберите файл")
                if length > 50_000_000:
                    raise ValueError("Файл слишком большой")
                raw = self.rfile.read(length)
                with connect() as conn:
                    if not conn.execute("SELECT 1 FROM projects WHERE id=?", (project_id,)).fetchone():
                        self.send_error_json("Стройка не найдена", 404)
                        return
                    if act_id is not None:
                        act = conn.execute(
                            "SELECT a.id FROM contract_acts a JOIN contracts c ON c.id=a.contract_id "
                            "WHERE a.id=? AND c.project_id=?",
                            (act_id, project_id),
                        ).fetchone()
                        if act is None:
                            self.send_error_json("Акт не найден", 404)
                            return
                    cur = conn.execute(
                        "INSERT INTO documents(project_id, act_id, filename, size) VALUES(?, ?, ?, ?)",
                        (project_id, act_id, filename, length),
                    )
                    doc_id = int(cur.lastrowid)
                    doc_dir = documents_dir(project_id)
                    doc_dir.mkdir(parents=True, exist_ok=True)
                    (doc_dir / f"{doc_id}_{filename}").write_bytes(raw)
                    state = get_state(conn)
                self.send_json({"id": doc_id, "state": state}, 201)
                return
            if path == "/api/import.xlsx":
                length = int(self.headers.get("Content-Length", "0") or 0)
                if length <= 0:
                    raise ValueError("Выберите файл Excel")
                if length > 25_000_000:
                    raise ValueError("Файл Excel слишком большой")
                raw = self.rfile.read(length)
                imported = parse_workbook(raw)
                with connect() as conn:
                    restore_state(conn, imported, reset=True)
                    state = get_state(conn)
                self.send_json({
                    "state": state,
                    "summary": {
                        "projects": len(imported.get("projects", [])),
                        "project_entries": sum(len(p.get("entries", [])) for p in imported.get("projects", [])),
                        "overheads": len(imported.get("overheads", [])),
                    },
                })
                return
            data = self.read_json()
            with connect() as conn:
                if path == "/api/projects":
                    name = clean_name(data.get("name"))
                    exists = conn.execute("SELECT 1 FROM projects WHERE lower(name)=lower(?)", (name,)).fetchone()
                    if exists:
                        raise ValueError("Стройка с таким названием уже существует")
                    order = conn.execute("SELECT COALESCE(MAX(sort_order),0)+1 FROM projects").fetchone()[0]
                    cur = conn.execute("INSERT INTO projects(name, sort_order) VALUES(?, ?)", (name, order))
                    self.send_json({"id": cur.lastrowid, "state": get_state(conn)}, 201)
                    return
                if path.startswith("/api/projects/") and path.endswith("/entries"):
                    project_id = int(path.split("/")[3])
                    if not conn.execute("SELECT 1 FROM projects WHERE id=?", (project_id,)).fetchone():
                        self.send_error_json("Стройка не найдена", 404)
                        return
                    section = str(data.get("section", ""))
                    if section not in PROJECT_SECTIONS:
                        raise ValueError("Некорректный раздел")
                    cur = conn.execute(
                        """INSERT INTO project_entries(project_id, section, category, description, amount, date, note)
                           VALUES(?, ?, ?, ?, ?, ?, ?)""",
                        (
                            project_id,
                            section,
                            str(data.get("category", "")).strip(),
                            str(data.get("description", "")).strip(),
                            clean_amount(data.get("amount")),
                            clean_date(data.get("date")),
                            str(data.get("note", "")).strip(),
                        ),
                    )
                    self.send_json({"id": cur.lastrowid, "state": get_state(conn)}, 201)
                    return
                if path == "/api/contracts":
                    project_id = int(data.get("project_id") or 0)
                    if not conn.execute("SELECT 1 FROM projects WHERE id=?", (project_id,)).fetchone():
                        self.send_error_json("Стройка не найдена", 404)
                        return
                    cur = conn.execute(
                        """INSERT INTO contracts(project_id, name, date, amount)
                           VALUES(?, ?, ?, ?)""",
                        (
                            project_id,
                            clean_text(data.get("name"), "Название договора не может быть пустым"),
                            clean_date(data.get("date")),
                            clean_amount(data.get("amount")),
                        ),
                    )
                    self.send_json({"id": cur.lastrowid, "state": get_state(conn)}, 201)
                    return
                if path.startswith("/api/contracts/") and path.endswith("/acts"):
                    contract_id = int(path.split("/")[3])
                    if not conn.execute("SELECT 1 FROM contracts WHERE id=?", (contract_id,)).fetchone():
                        self.send_error_json("Договор не найден", 404)
                        return
                    cur = conn.execute(
                        """INSERT INTO contract_acts(contract_id, number, date, amount)
                           VALUES(?, ?, ?, ?)""",
                        (
                            contract_id,
                            clean_text(data.get("number"), "Номер акта не может быть пустым"),
                            clean_date(data.get("date")),
                            clean_amount(data.get("amount")),
                        ),
                    )
                    self.send_json({"id": cur.lastrowid, "state": get_state(conn)}, 201)
                    return
                if path == "/api/overheads":
                    group = str(data.get("group_type", "cashless"))
                    if group not in OVERHEAD_GROUPS:
                        raise ValueError("Некорректный вид накладных")
                    cur = conn.execute(
                        """INSERT INTO overhead_entries(group_type, category, description, amount, date)
                           VALUES(?, ?, ?, ?, ?)""",
                        (
                            group,
                            str(data.get("category", "")).strip(),
                            str(data.get("description", "")).strip(),
                            clean_amount(data.get("amount")),
                            clean_date(data.get("date")),
                        ),
                    )
                    self.send_json({"id": cur.lastrowid, "state": get_state(conn)}, 201)
                    return
                if path == "/api/settings":
                    rate = float(data.get("conversion_rate"))
                    if not 0 <= rate <= 1:
                        raise ValueError("Ставка должна быть от 0 до 1")
                    conn.execute(
                        "INSERT INTO settings(key,value) VALUES('conversion_rate',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                        (str(rate),),
                    )
                    self.send_json({"state": get_state(conn)})
                    return
                if path == "/api/restore":
                    restore_state(conn, data, reset=True)
                    self.send_json({"state": get_state(conn)})
                    return
            self.send_error_json("Маршрут не найден", 404)
        except ValueError as exc:
            self.send_error_json(str(exc), 400)
        except Exception as exc:
            self.send_error_json(str(exc), 500)

    def do_PUT(self) -> None:
        path = unquote(urlparse(self.path).path)
        try:
            data = self.read_json()
            with connect() as conn:
                parts = [p for p in path.split("/") if p]
                if len(parts) == 3 and parts[:2] == ["api", "projects"]:
                    project_id = int(parts[2])
                    name = clean_name(data.get("name"))
                    exists = conn.execute("SELECT 1 FROM projects WHERE lower(name)=lower(?) AND id<>?", (name, project_id)).fetchone()
                    if exists:
                        raise ValueError("Стройка с таким названием уже существует")
                    cur = conn.execute("UPDATE projects SET name=? WHERE id=?", (name, project_id))
                    if not cur.rowcount:
                        self.send_error_json("Стройка не найдена", 404)
                        return
                    self.send_json({"state": get_state(conn)})
                    return
                if len(parts) == 3 and parts[:2] == ["api", "contracts"]:
                    contract_id = int(parts[2])
                    cur = conn.execute(
                        "UPDATE contracts SET name=?, date=?, amount=? WHERE id=?",
                        (
                            clean_text(data.get("name"), "Название договора не может быть пустым"),
                            clean_date(data.get("date")),
                            clean_amount(data.get("amount")),
                            contract_id,
                        ),
                    )
                    if not cur.rowcount:
                        self.send_error_json("Договор не найден", 404)
                        return
                    self.send_json({"state": get_state(conn)})
                    return
                if len(parts) == 3 and parts[:2] == ["api", "acts"]:
                    act_id = int(parts[2])
                    cur = conn.execute(
                        "UPDATE contract_acts SET number=?, date=?, amount=? WHERE id=?",
                        (
                            clean_text(data.get("number"), "Номер акта не может быть пустым"),
                            clean_date(data.get("date")),
                            clean_amount(data.get("amount")),
                            act_id,
                        ),
                    )
                    if not cur.rowcount:
                        self.send_error_json("Акт не найден", 404)
                        return
                    self.send_json({"state": get_state(conn)})
                    return
                if len(parts) == 3 and parts[:2] == ["api", "entries"]:
                    entry_id = int(parts[2])
                    section = str(data.get("section", ""))
                    if section not in PROJECT_SECTIONS:
                        raise ValueError("Некорректный раздел")
                    cur = conn.execute(
                        """UPDATE project_entries SET section=?, category=?, description=?, amount=?, date=?, note=? WHERE id=?""",
                        (
                            section,
                            str(data.get("category", "")).strip(),
                            str(data.get("description", "")).strip(),
                            clean_amount(data.get("amount")),
                            clean_date(data.get("date")),
                            str(data.get("note", "")).strip(),
                            entry_id,
                        ),
                    )
                    if not cur.rowcount:
                        self.send_error_json("Запись не найдена", 404)
                        return
                    self.send_json({"state": get_state(conn)})
                    return
                if len(parts) == 3 and parts[:2] == ["api", "overheads"]:
                    entry_id = int(parts[2])
                    group = str(data.get("group_type", "cashless"))
                    if group not in OVERHEAD_GROUPS:
                        raise ValueError("Некорректный вид накладных")
                    cur = conn.execute(
                        """UPDATE overhead_entries SET group_type=?, category=?, description=?, amount=?, date=? WHERE id=?""",
                        (
                            group,
                            str(data.get("category", "")).strip(),
                            str(data.get("description", "")).strip(),
                            clean_amount(data.get("amount")),
                            clean_date(data.get("date")),
                            entry_id,
                        ),
                    )
                    if not cur.rowcount:
                        self.send_error_json("Запись не найдена", 404)
                        return
                    self.send_json({"state": get_state(conn)})
                    return
            self.send_error_json("Маршрут не найден", 404)
        except ValueError as exc:
            self.send_error_json(str(exc), 400)
        except Exception as exc:
            self.send_error_json(str(exc), 500)

    def do_DELETE(self) -> None:
        path = unquote(urlparse(self.path).path)
        try:
            parts = [p for p in path.split("/") if p]
            with connect() as conn:
                if len(parts) == 3 and parts[:2] == ["api", "contracts"]:
                    contract_id = int(parts[2])
                    if not conn.execute("SELECT 1 FROM contracts WHERE id=?", (contract_id,)).fetchone():
                        self.send_error_json("Договор не найден", 404)
                        return
                    act_ids = [r[0] for r in conn.execute(
                        "SELECT id FROM contract_acts WHERE contract_id=?", (contract_id,)
                    )]
                    remove_act_documents(conn, act_ids)
                    conn.execute("DELETE FROM contracts WHERE id=?", (contract_id,))
                    self.send_json({"state": get_state(conn)})
                    return
                if len(parts) == 3 and parts[:2] == ["api", "acts"]:
                    act_id = int(parts[2])
                    if not conn.execute("SELECT 1 FROM contract_acts WHERE id=?", (act_id,)).fetchone():
                        self.send_error_json("Акт не найден", 404)
                        return
                    remove_act_documents(conn, [act_id])
                    conn.execute("DELETE FROM contract_acts WHERE id=?", (act_id,))
                    self.send_json({"state": get_state(conn)})
                    return
                if len(parts) == 3 and parts[:2] == ["api", "projects"]:
                    project_id = int(parts[2])
                    cur = conn.execute("DELETE FROM projects WHERE id=?", (project_id,))
                    if not cur.rowcount:
                        self.send_error_json("Стройка не найдена", 404)
                        return
                    shutil.rmtree(documents_dir(project_id), ignore_errors=True)
                    self.send_json({"state": get_state(conn)})
                    return
                if len(parts) == 3 and parts[:2] == ["api", "documents"]:
                    doc_id = int(parts[2])
                    row = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
                    if row is None:
                        self.send_error_json("Файл не найден", 404)
                        return
                    remove_document_files(row)
                    conn.execute("DELETE FROM documents WHERE id=?", (doc_id,))
                    self.send_json({"state": get_state(conn)})
                    return
                if len(parts) == 3 and parts[:2] == ["api", "entries"]:
                    cur = conn.execute("DELETE FROM project_entries WHERE id=?", (int(parts[2]),))
                    if not cur.rowcount:
                        self.send_error_json("Запись не найдена", 404)
                        return
                    self.send_json({"state": get_state(conn)})
                    return
                if len(parts) == 3 and parts[:2] == ["api", "overheads"]:
                    cur = conn.execute("DELETE FROM overhead_entries WHERE id=?", (int(parts[2]),))
                    if not cur.rowcount:
                        self.send_error_json("Запись не найдена", 404)
                        return
                    self.send_json({"state": get_state(conn)})
                    return
            self.send_error_json("Маршрут не найден", 404)
        except Exception as exc:
            self.send_error_json(str(exc), 500)

    def serve_static(self, path: str) -> None:
        if path in {"", "/"}:
            target = STATIC_DIR / "index.html"
        else:
            relative = Path(path.lstrip("/"))
            if ".." in relative.parts:
                self.send_error(HTTPStatus.FORBIDDEN)
                return
            target = STATIC_DIR / relative
        if not target.exists() or not target.is_file():
            target = STATIC_DIR / "index.html"
        payload = target.read_bytes()
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        if content_type.startswith("text/") or content_type in {"application/javascript", "application/json"}:
            content_type += "; charset=utf-8"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(payload)


def find_port(start: int = 8765, attempts: int = 20) -> int:
    import socket
    for port in range(start, start + attempts):
        with socket.socket() as sock:
            try:
                sock.bind((HOST, port))
            except OSError:
                continue
            return port
    raise RuntimeError("Не удалось найти свободный порт")


def main() -> None:
    parser = argparse.ArgumentParser(description="СтройФинанс — учёт строительных проектов")
    parser.add_argument("--port", type=int, default=0)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    init_db()
    port = args.port or find_port()
    server = ThreadingHTTPServer((HOST, port), Handler)
    url = f"http://{HOST}:{port}"
    print(f"СтройФинанс запущен: {url}")
    print("Чтобы остановить приложение, закройте это окно или нажмите Ctrl+C.")
    if not args.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
