"""Runs an importer for an uploaded file and saves its rows into SQLite (replacing that file's earlier rows)."""
import threading
import traceback
from pathlib import Path

from . import ageing, batch, orders, reasons
from .specs import KIND_TABLES, TABLES

PARSERS = {"ageing": ageing.parse, "orders": orders.parse, "reasons": reasons.parse, "template": batch.parse, "validate": batch.parse}
SERVER_PARSED = set(PARSERS)          # ar_pdf is parsed in the browser (pdf.js)


def save_tables(conn, file_id: int, filename: str, tables: dict) -> int:
    from .. import data_cache
    data_cache.invalidate(file_id)          # the cached download is stale as soon as the rows change
    total = 0
    for tid, rows in tables.items():
        cols = ['"%s"' % c["key"] for c in TABLES[tid]["cols"] if not c.get("virt")]
        conn.execute(f"DELETE FROM t_{tid} WHERE source_file_id = ?", (file_id,))
        q = f"INSERT INTO t_{tid} (source_file_id, source_file, row_no, {', '.join(cols)}) VALUES (?,?,?,{','.join('?' * len(cols))})"
        conn.executemany(q, ((file_id, filename, i + 1, *r) for i, r in enumerate(rows)))
        total += len(rows)
    return total


def set_status(conn, file_id: int, status: str, note: str = "", rows: int = 0) -> None:
    conn.execute("UPDATE source_files SET status = ?, note = ?, rows_imported = ? WHERE id = ?", (status, note, rows, file_id))


def import_file(file_id: int) -> None:
    """Blocking. Called from a background thread."""
    from ..database import ROOT, transaction
    with transaction() as conn:
        f = conn.execute("SELECT * FROM source_files WHERE id = ?", (file_id,)).fetchone()
        if not f or f["kind"] not in SERVER_PARSED:
            return
        set_status(conn, file_id, "importing", "Starting")
    path = Path(f["stored_path"])
    path = path if path.is_absolute() else ROOT / path

    def progress(msg: str) -> None:
        with transaction() as c:
            c.execute("UPDATE source_files SET note = ? WHERE id = ?", (msg, file_id))

    try:
        tables = PARSERS[f["kind"]](path, progress)
        with transaction() as conn:
            n = save_tables(conn, file_id, f["filename"], tables)
            set_status(conn, file_id, "imported", " · ".join(f"{len(r):,} {TABLES[t]['title']}" for t, r in tables.items() if r), n)
        from .. import data_cache
        data_cache.build(file_id)           # prepare the gzip download now, so the first view is instant
    except Exception as e:  # report the reason to the user instead of failing silently
        traceback.print_exc()
        with transaction() as conn:
            set_status(conn, file_id, "failed", str(e)[:300])


def start_import(file_id: int) -> None:
    threading.Thread(target=import_file, args=(file_id,), daemon=True).start()


def tables_for_kind(kind: str) -> list:
    return KIND_TABLES.get(kind, [])
