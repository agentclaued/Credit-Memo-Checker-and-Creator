"""SQLite connection and schema. One local file: data/credit.sqlite."""
import os
import sqlite3
from contextlib import contextmanager
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DB_PATH = Path(os.environ.get("CREDIT_DB_PATH", ROOT / "data" / "credit.sqlite"))
UPLOAD_DIR = Path(os.environ.get("CREDIT_UPLOAD_DIR", ROOT / "data" / "uploads"))

SCHEMA = """
CREATE TABLE IF NOT EXISTS source_files (
    id INTEGER PRIMARY KEY,
    kind TEXT NOT NULL,            -- ageing | orders | ar_pdf | template | validate
    filename TEXT NOT NULL,        -- original file name (kept for every imported row later)
    stored_path TEXT NOT NULL,
    size_bytes INTEGER NOT NULL,
    status TEXT NOT NULL DEFAULT 'uploaded',   -- uploaded | imported | failed
    rows_imported INTEGER NOT NULL DEFAULT 0,
    note TEXT NOT NULL DEFAULT '',
    uploaded_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
"""


def connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_schema() -> None:
    with transaction() as conn:
        conn.executescript(SCHEMA)
        from .importers.specs import schema_sql
        conn.executescript(schema_sql())
        from .importers.specs import SQL_TYPE, TABLES
        for tid, t in TABLES.items():        # tables made before a column was added: add it, never drop or rewrite data
            have = {r["name"] for r in conn.execute(f"PRAGMA table_info(t_{tid})")}
            for c in t["cols"]:
                if not c.get("virt") and c["key"] not in have:
                    conn.execute(f'ALTER TABLE t_{tid} ADD COLUMN "{c["key"]}" {SQL_TYPE[c["type"]]}')
        conn.execute("UPDATE t_orders SET dv_sale_report = COALESCE(dv,'') || COALESCE(sale_report,'') "
                     "WHERE dv_sale_report IS NULL AND (dv IS NOT NULL OR sale_report IS NOT NULL)")   # fill rows imported before the column existed


@contextmanager
def transaction():
    conn = connect()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
