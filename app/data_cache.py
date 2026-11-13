"""Fast data pipeline for the viewer.

* payload(file_id, limit): the tables of one stored file as plain arrays (limit = first N rows per table, for the instant first screen).
* get(file_id): the FULL payload as a ready-made gzip file (built once, after the import) so a request is just a file read.
"""
import gzip
import json
import threading

from .database import DB_PATH, transaction
from .importers.specs import KIND_TABLES, TABLES

CACHE_DIR = DB_PATH.parent / "cache"
_lock = threading.Lock()


def payload(file_id: int, limit: int | None = None) -> dict | None:
    with transaction() as conn:
        f = conn.execute("SELECT * FROM source_files WHERE id = ?", (file_id,)).fetchone()
        if not f:
            return None
        tables = []
        for tid in KIND_TABLES.get(f["kind"], []):
            cols = TABLES[tid]["cols"]
            select = f"SELECT {', '.join(chr(34) + c['key'] + chr(34) for c in cols)} FROM t_{tid} WHERE source_file_id = ? ORDER BY row_no"
            args = (file_id,)
            if limit:
                select += " LIMIT ?"
                args += (limit,)
            rows = [tuple(r) for r in conn.execute(select, args)]
            total = conn.execute(f"SELECT COUNT(*) FROM t_{tid} WHERE source_file_id = ?", (file_id,)).fetchone()[0]
            tables.append({"id": tid, "title": TABLES[tid]["title"], "cols": cols, "total": total, "rows": rows})
        return {"file": dict(f), "tables": tables, "partial": bool(limit)}


def _signature(file_id: int) -> str:
    """Column layout of the file's tables: a cache built before a column was added/moved is stale."""
    with transaction() as conn:
        f = conn.execute("SELECT kind FROM source_files WHERE id = ?", (file_id,)).fetchone()
    tids = KIND_TABLES.get(f["kind"], []) if f else []
    return "|".join(c["key"] for tid in tids for c in TABLES[tid]["cols"])


def _paths(file_id: int):
    return CACHE_DIR / f"file_{file_id}.json.gz", CACHE_DIR / f"file_{file_id}.meta"


def invalidate(file_id: int) -> None:
    with _lock:
        for p in _paths(file_id):
            p.unlink(missing_ok=True)


def build(file_id: int) -> bool:
    data = payload(file_id)
    if data is None:
        return False
    raw = json.dumps(data, separators=(",", ":")).encode("utf-8")
    gz = gzip.compress(raw, 4)
    with _lock:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        data_path, meta_path = _paths(file_id)
        data_path.write_bytes(gz)
        meta_path.write_text(json.dumps({"raw": len(raw), "sig": _signature(file_id)}))
    return True


def get(file_id: int):
    """(gzip bytes, uncompressed size, etag) of the full payload, building it first if needed. None if the file is unknown."""
    data_path, meta_path = _paths(file_id)
    meta = json.loads(meta_path.read_text()) if meta_path.exists() else {}
    if not data_path.exists() or meta.get("sig") != _signature(file_id):
        if not build(file_id):
            return None
        meta = json.loads(meta_path.read_text())
    st = data_path.stat()
    raw = meta["raw"]
    return data_path.read_bytes(), raw, f'"{file_id}-{st.st_mtime_ns}-{st.st_size}"'
