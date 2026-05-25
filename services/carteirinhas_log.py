import json
import os
import sqlite3
from contextlib import closing
from datetime import datetime


DEFAULT_BASE_DIR = "uploads"
DEFAULT_JSON_NAME = "carteirinhas_print_log.json"
DEFAULT_DB_NAME = "carteirinhas_print_log.sqlite3"


def printlog_json_path(base_dir=DEFAULT_BASE_DIR):
    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, DEFAULT_JSON_NAME)


def printlog_db_path(base_dir=DEFAULT_BASE_DIR):
    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, DEFAULT_DB_NAME)


def _init_print_log_db(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS printed_cards (
            ano INTEGER NOT NULL,
            rm INTEGER NOT NULL,
            printed_at TEXT NOT NULL,
            PRIMARY KEY (ano, rm)
        )
        """
    )


def _migrate_print_log_json(conn, json_path):
    if not json_path or not os.path.exists(json_path):
        return

    try:
        cur = conn.execute("SELECT COUNT(*) FROM printed_cards")
        if cur.fetchone()[0] > 0:
            return
    except Exception:
        return

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f) or {}
    except Exception:
        return

    for ano_raw, bloco in data.items():
        try:
            ano = int(ano_raw)
        except Exception:
            continue

        printed_at = bloco.get("printed_at", {}) if isinstance(bloco, dict) else {}
        printed_rms = bloco.get("printed_rms", []) if isinstance(bloco, dict) else []
        for rm_raw in printed_rms:
            try:
                rm = int(rm_raw)
            except Exception:
                continue
            ts = str(printed_at.get(str(rm), "")).strip() or datetime.now().isoformat(timespec="seconds")
            conn.execute(
                "INSERT OR IGNORE INTO printed_cards (ano, rm, printed_at) VALUES (?, ?, ?)",
                (ano, rm, ts),
            )


def connect_print_log_db(db_path=None, json_path=None, base_dir=DEFAULT_BASE_DIR):
    db_path = printlog_db_path(base_dir) if db_path is None else db_path
    json_path = printlog_json_path(base_dir) if json_path is None else json_path
    os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)

    conn = sqlite3.connect(db_path, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=30000")
    _init_print_log_db(conn)
    _migrate_print_log_json(conn, json_path)
    conn.commit()
    return conn


def load_print_log(db_path=None, json_path=None, base_dir=DEFAULT_BASE_DIR):
    log = {}
    with closing(connect_print_log_db(db_path=db_path, json_path=json_path, base_dir=base_dir)) as conn:
        rows = conn.execute("SELECT ano, rm, printed_at FROM printed_cards ORDER BY ano, rm").fetchall()

    for ano, rm, printed_at in rows:
        key = str(ano)
        log.setdefault(key, {"printed_rms": [], "printed_at": {}})
        log[key]["printed_rms"].append(rm)
        log[key]["printed_at"][str(rm)] = printed_at
    return log


def save_print_log(data: dict, db_path=None, json_path=None, base_dir=DEFAULT_BASE_DIR):
    with closing(connect_print_log_db(db_path=db_path, json_path=json_path, base_dir=base_dir)) as conn:
        conn.execute("BEGIN IMMEDIATE")
        conn.execute("DELETE FROM printed_cards")
        for ano_raw, bloco in (data or {}).items():
            try:
                ano = int(ano_raw)
            except Exception:
                continue

            printed_at = bloco.get("printed_at", {}) if isinstance(bloco, dict) else {}
            printed_rms = bloco.get("printed_rms", []) if isinstance(bloco, dict) else []
            for rm_raw in printed_rms:
                try:
                    rm = int(rm_raw)
                except Exception:
                    continue
                ts = str(printed_at.get(str(rm), "")).strip() or datetime.now().isoformat(timespec="seconds")
                conn.execute(
                    "INSERT OR REPLACE INTO printed_cards (ano, rm, printed_at) VALUES (?, ?, ?)",
                    (ano, rm, ts),
                )
        conn.commit()


def get_printed_set(ano: int, db_path=None, json_path=None, base_dir=DEFAULT_BASE_DIR) -> set:
    with closing(connect_print_log_db(db_path=db_path, json_path=json_path, base_dir=base_dir)) as conn:
        rows = conn.execute("SELECT rm FROM printed_cards WHERE ano = ?", (int(ano),)).fetchall()
    return {int(row[0]) for row in rows}


def mark_printed_rms(ano: int, rms: list[int], db_path=None, json_path=None, base_dir=DEFAULT_BASE_DIR) -> tuple[int, int]:
    now_iso = datetime.now().isoformat(timespec="seconds")
    with closing(connect_print_log_db(db_path=db_path, json_path=json_path, base_dir=base_dir)) as conn:
        conn.execute("BEGIN IMMEDIATE")
        before = {
            int(row[0])
            for row in conn.execute("SELECT rm FROM printed_cards WHERE ano = ?", (int(ano),)).fetchall()
        }
        for rm in rms:
            conn.execute(
                "INSERT OR REPLACE INTO printed_cards (ano, rm, printed_at) VALUES (?, ?, ?)",
                (int(ano), int(rm), now_iso),
            )
        after = {
            int(row[0])
            for row in conn.execute("SELECT rm FROM printed_cards WHERE ano = ?", (int(ano),)).fetchall()
        }
        conn.commit()
    return len(after - before), len(after)
