"""Credit Department system — FastAPI routes (pages, JSON API, upload handler)."""
import re
import shutil
import tempfile
import threading
from datetime import datetime
from functools import lru_cache
from pathlib import Path

from contextlib import asynccontextmanager

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from python_calamine import load_workbook

from . import autofill, data_cache, exporter, recheck, template_rules
from .blocked_reasons import is_blocked
from .database import ROOT, UPLOAD_DIR, init_schema, transaction
from .importers import SOURCE_BY_KEY, SOURCES, ar_pdf, runner
from .importers.specs import KIND_TABLES, TABLES

APP_VERSION = "0.1.0-demo"
BUILD_TIME = datetime.now().strftime("%Y%m%d%H%M%S")  # cache-busting only
TEMPLATE_PATH = ROOT / "templates_xlsx" / template_rules.TEMPLATE_FILE
MAX_UPLOAD_BYTES = 200 * 1024 * 1024

@asynccontextmanager
async def lifespan(_app: FastAPI):
    init_schema()
    with transaction() as conn:   # finish imports that were interrupted or uploaded before importing existed
        todo = conn.execute("SELECT id FROM source_files WHERE status IN ('uploaded','importing') AND kind IN (%s)"
                            % ",".join("?" * len(runner.SERVER_PARSED)), tuple(runner.SERVER_PARSED)).fetchall()
    for r in todo:
        runner.start_import(r["id"])
    yield


app = FastAPI(title="Credit Department", version=APP_VERSION, lifespan=lifespan)
app.add_middleware(GZipMiddleware, minimum_size=1000)
app.mount("/static", StaticFiles(directory=Path(__file__).parent / "static"), name="static")
templates = Jinja2Templates(directory=Path(__file__).parent / "templates")


def render(request: Request, name: str, **ctx) -> HTMLResponse:
    ctx.update(build_time=BUILD_TIME, version=APP_VERSION, sources=SOURCES, path=request.url.path)
    return templates.TemplateResponse(request, name, ctx)


# ----------------------------------------------------------------------------- pages
@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request):
    return RedirectResponse("/recheck")      # the dashboard is hidden; Recheck is the main page


@app.get("/upload", response_class=HTMLResponse)
def upload_page(request: Request):
    return render(request, "upload.html")


@app.get("/memo", response_class=HTMLResponse)
def memo_page(request: Request):
    return render(request, "memo.html", rules=template_rules)


@app.get("/recheck", response_class=HTMLResponse)
def recheck_page(request: Request):
    return render(request, "recheck.html", checks=[(recheck.TAGS.get(n, f'P{n}') if n in (19, 32) else f'P{n}', t, i) for n, t, i in recheck.CHECKS])


@app.get("/rules", response_class=HTMLResponse)
def rules_page(request: Request):
    return render(request, "rules.html", rules=template_rules)


# ----------------------------------------------------------------------------- upload
def safe_name(name: str) -> str:
    return re.sub(r"[^\w.\- ()#$,]+", "_", Path(name).name).strip() or "file"


@app.post("/upload/{kind}")
async def upload_file(kind: str, file: UploadFile = File(...)):
    source = SOURCE_BY_KEY.get(kind)
    if not source:
        raise HTTPException(404, "Unknown file type.")
    ext = Path(file.filename or "").suffix.lower()
    if ext not in source["accept"].split(","):
        raise HTTPException(400, f"{source['label']} accepts {source['accept']} files only.")
    data = await file.read()
    if not data:
        raise HTTPException(400, "The file is empty.")
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(400, "The file is larger than 200 MB.")
    with transaction() as conn:       # the same file name twice in the same section
        dup = conn.execute("SELECT * FROM source_files WHERE kind = ? AND lower(filename) = lower(?) AND status != 'failed' ORDER BY id DESC",
                           (kind, file.filename)).fetchone()
    folder = UPLOAD_DIR / kind
    folder.mkdir(parents=True, exist_ok=True)
    stored = folder / f"{datetime.now():%Y%m%d-%H%M%S}-{safe_name(file.filename)}"
    if dup:
        old = Path(dup["stored_path"])
        old = old if old.is_absolute() else ROOT / old
        if old.exists() and old.read_bytes() == data:        # identical file: nothing to update
            raise HTTPException(409, f"\"{file.filename}\" is already uploaded in {source['label']} ({dup['uploaded_at']}) and has no changes.")
        # same name, changed content: keep the old copy on disk and update this file's rows from the new one
        stored.write_bytes(data)
        with transaction() as conn:
            conn.execute("UPDATE source_files SET stored_path = ?, size_bytes = ?, status = 'uploaded', note = '', uploaded_at = datetime('now','localtime') WHERE id = ?",
                         (str(stored), len(data), dup["id"]))
        if kind in runner.SERVER_PARSED:
            runner.start_import(dup["id"])                   # replaces only this file's rows
        return {"id": dup["id"], "kind": kind, "filename": file.filename, "size": len(data), "updated": True}
    stored.write_bytes(data)
    with transaction() as conn:
        cur = conn.execute(
            "INSERT INTO source_files (kind, filename, stored_path, size_bytes) VALUES (?,?,?,?)",
            (kind, file.filename, str(stored), len(data)),
        )
    if kind in runner.SERVER_PARSED:
        runner.start_import(cur.lastrowid)       # pdf is read in the browser (pdf.js) and posted to /api/files/{id}/rows
    return {"id": cur.lastrowid, "kind": kind, "filename": file.filename, "size": len(data)}


@app.post("/api/recheck")
async def recheck_batch(file: UploadFile = File(...)):
    """Run the 10 recheck points on an uploaded batch workbook. The file is not stored."""
    if Path(file.filename or "").suffix.lower() != ".xlsx":
        raise HTTPException(400, "Upload the credit memo batch as an .xlsx file.")
    data = await file.read()
    if not data:
        raise HTTPException(400, "The file is empty.")
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(400, "The file is larger than 200 MB.")
    with tempfile.TemporaryDirectory() as tmp:      # temporary copy, deleted right after the check
        stored = Path(tmp) / safe_name(file.filename)
        stored.write_bytes(data)
        try:
            result = recheck.run(stored)
        except Exception as e:
            raise HTTPException(400, f"This is not a readable Credit Memo Batch file ({e}).")
    if not result["memos"]:
        raise HTTPException(400, "No filled memo sheets found in this file.")
    return {"filename": file.filename, **result}


@app.get("/api/files")
def list_files():
    with transaction() as conn:
        rows = conn.execute("SELECT * FROM source_files ORDER BY id DESC").fetchall()
    return [dict(r) for r in rows]


def stored_file(file_id: int):
    with transaction() as conn:
        f = conn.execute("SELECT * FROM source_files WHERE id = ?", (file_id,)).fetchone()
    if not f:
        raise HTTPException(404, "File not found.")
    path = Path(f["stored_path"])
    return f, (path if path.is_absolute() else ROOT / path)


@app.get("/api/files/{file_id}/download")
def download_file(file_id: int):
    f, path = stored_file(file_id)
    return FileResponse(path, filename=f["filename"])


@app.delete("/api/files/{file_id}")
def delete_file(file_id: int):
    """Remove the file's rows from every table and its upload record. The stored copy is moved to data/uploads/_deleted (recoverable)."""
    f, path = stored_file(file_id)
    if f["status"] == "importing":
        raise HTTPException(409, "This file is being imported. Try again in a moment.")
    with transaction() as conn:
        removed = 0
        for tid in TABLES:
            removed += conn.execute(f"DELETE FROM t_{tid} WHERE source_file_id = ?", (file_id,)).rowcount
        conn.execute("DELETE FROM source_files WHERE id = ?", (file_id,))
    data_cache.invalidate(file_id)
    kept = None
    if path.exists():
        dest_dir = UPLOAD_DIR / "_deleted"
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / f"{datetime.now():%Y%m%d-%H%M%S}-{path.name}"
        shutil.move(str(path), str(dest))
        kept = str(dest)
    return {"deleted_rows": removed, "kept_copy": kept}


@app.post("/api/files/{file_id}/import")
def reimport(file_id: int):
    f, _ = stored_file(file_id)
    if f["kind"] not in runner.SERVER_PARSED:
        raise HTTPException(400, "This file type is read in the browser.")
    runner.start_import(file_id)
    return {"ok": True}


@app.post("/api/files/{file_id}/rows")
async def receive_pdf_rows(file_id: int, request: Request):
    """The browser (pdf.js + Report2Excel parser) posts the parsed A/R tables of a PDF here."""
    f, _ = stored_file(file_id)
    if f["kind"] != "ar_pdf":
        raise HTTPException(400, "Only A/R PDF files are parsed in the browser.")
    payload = await request.json()
    try:
        tables = ar_pdf.rows_from_payload(payload)
    except ValueError as e:
        with transaction() as conn:
            runner.set_status(conn, file_id, "failed", str(e))
        raise HTTPException(400, str(e))
    warns = payload.get("warnings") or []
    with transaction() as conn:
        n = runner.save_tables(conn, file_id, f["filename"], tables)
        note = " · ".join(f"{len(r):,} {TABLES[t]['title']}" for t, r in tables.items() if r)
        runner.set_status(conn, file_id, "imported", note + (f" · {len(warns)} warning(s): {warns[0]}" if warns else ""), n)
    threading.Thread(target=data_cache.build, args=(file_id,), daemon=True).start()
    return {"rows": n}


@app.get("/api/files/{file_id}/reasons")
def get_reasons(file_id: int):
    f, _ = stored_file(file_id)
    if f["kind"] != "reasons":
        raise HTTPException(400, "Only the reason code list can be edited.")
    with transaction() as conn:
        rows = conn.execute("SELECT code, description FROM t_reason_list WHERE source_file_id = ? ORDER BY row_no", (file_id,)).fetchall()
    return [dict(r) for r in rows]


@app.put("/api/files/{file_id}/reasons")
async def save_reasons(file_id: int, request: Request):
    """Replace the reason codes of this list with the edited rows (the uploaded file itself is not changed)."""
    f, _ = stored_file(file_id)
    if f["kind"] != "reasons":
        raise HTTPException(400, "Only the reason code list can be edited.")
    payload = await request.json()
    rows, seen = [], set()
    for r in payload if isinstance(payload, list) else []:
        from .importers.reasons import fix_code, fix_desc
        code, desc = fix_code(r.get("code")), fix_desc(r.get("description")) or ""
        if not code and not desc:
            continue
        if not code:
            raise HTTPException(400, f"A description has no code: {desc}")
        if code in seen:
            raise HTTPException(400, f"Code {code} is in the list twice.")
        seen.add(code)
        rows.append([code, desc or None])
    if not rows:
        raise HTTPException(400, "The list cannot be empty.")
    with transaction() as conn:
        n = runner.save_tables(conn, file_id, f["filename"], {"reason_list": rows})
        runner.set_status(conn, file_id, "imported", f"{n:,} Reason codes · edited by hand", n)
    data_cache.build(file_id)
    return {"rows": n}


@app.get("/api/files/{file_id}/data")
def file_data(file_id: int, request: Request, limit: int | None = None):
    """Tables of one uploaded file as arrays in column order.
    ?limit=N -> only the first N rows per table (instant first screen, with the real totals).
    no limit -> the full data from a ready-made gzip cache (built after the import; ETag makes repeat views a 304)."""
    stored_file(file_id)
    if limit:
        data = data_cache.payload(file_id, min(max(limit, 1), 1000))
        return JSONResponse(data)
    got = data_cache.get(file_id)
    if got is None:
        raise HTTPException(404, "File not found.")
    body, raw_size, etag = got
    headers = {"ETag": etag, "Cache-Control": "no-cache", "X-Raw-Size": str(raw_size), "Access-Control-Expose-Headers": "X-Raw-Size"}
    if request.headers.get("if-none-match") == etag:
        return Response(status_code=304, headers=headers)
    return Response(body, media_type="application/json", headers={**headers, "Content-Encoding": "gzip"})


def stored_counts() -> list:
    """Rows currently held in each SQLite table (all files)."""
    out = []
    with transaction() as conn:
        for tid, t in TABLES.items():
            n = conn.execute(f"SELECT COUNT(*) FROM t_{tid}").fetchone()[0]
            out.append({"id": tid, "title": t["title"], "kind": t["kind"], "rows": n})
    return out


def file_counts() -> dict:
    with transaction() as conn:
        rows = conn.execute("SELECT kind, COUNT(*) n, MAX(uploaded_at) last FROM source_files GROUP BY kind").fetchall()
    return {r["kind"]: {"n": r["n"], "last": r["last"]} for r in rows}


# ----------------------------------------------------------------------------- template data (read from the master template)
# The template's PRICE LIST is empty; the demo falls back to the price list of the filled sample batch.
SAMPLE_BATCH = ROOT / "Rechecking_process" / "RAW Data" / "input file to validate its data" / "Credit Memo Batch #10-26973 $13,099.16.xlsx"


@lru_cache(maxsize=1)
def template_sheet_rows() -> dict:
    wb = load_workbook(str(TEMPLATE_PATH))
    out = {n: wb.get_sheet_by_name(n).to_python() for n in ("REASON", "PRICE LIST")}
    if len(out["PRICE LIST"]) < 2 and SAMPLE_BATCH.exists():
        out["PRICE LIST"] = load_workbook(str(SAMPLE_BATCH)).get_sheet_by_name("PRICE LIST").to_python()
    return out


@app.get("/api/price-list")
def price_list():
    rows = template_sheet_rows()["PRICE LIST"][1:]
    return [{"item": str(r[0]).strip(), "description": str(r[1]).strip()} for r in rows if r and r[0] not in ("", None)]


@app.get("/api/reasons")
def reasons():
    rows = template_sheet_rows()["REASON"][2:]
    return [{"code": str(r[0]).strip(), "description": str(r[1]).strip(), "blocked": is_blocked(str(r[0]), str(r[1]))}
            for r in rows if r and r[0] not in ("", None)]


@app.post("/api/autofill")
async def autofill_memos(request: Request):
    """Memo details (account, name, salesman, terms, ship code) for the DB#/IN# numbers pasted into the Batch sheet."""
    data = await request.json()
    return autofill.lookup([str(x) for x in (data.get("refs") or [])][:1000], str(data.get("division") or ""))


@app.post("/api/export")
async def export_batch(request: Request):
    """Creates the Excel file from the Create-batch draft: a copy of the template, filled in its own cells only."""
    data = await request.json()
    rows = template_sheet_rows()
    reasons_map = {str(r[0]).strip(): str(r[1]).strip() for r in rows["REASON"][2:] if r and r[0] not in ("", None)}
    for i, m in enumerate(data.get("memos") or [], 1):
        code = str(m.get("code") or "").strip()
        if is_blocked(code, reasons_map.get(code, "")):
            raise HTTPException(400, f"Memo {i}: reason code {code} ({reasons_map.get(code, '')}) is not allowed for a credit memo.")
    try:
        buf = exporter.build(data, TEMPLATE_PATH, rows["PRICE LIST"], reasons_map)
    except ValueError as e:
        raise HTTPException(400, str(e))
    name = re.sub(r'[\/:*?"<>|]', "", f"Credit Memo Batch #{(data.get('batch') or {}).get('no') or 'draft'}") + ".xlsx"
    return Response(buf.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})


@app.get("/api/health")
def health():
    return JSONResponse({"ok": True, "version": APP_VERSION})
