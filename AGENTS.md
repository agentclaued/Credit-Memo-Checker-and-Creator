# Agent Instructions

## Project

Native FastAPI app for the Credit Department: import raw files into local SQLite, create credit memo batches in a web UI, export them as Excel in the user's template, and recheck filled batches. Entry point: `app/main.py`.

## Working rules

- Keep the application native FastAPI; do not add Flask or WSGI wrappers.
- STRICT TEMPLATE: exports and validation may use only the sheets, columns, formulas and formatting of `templates_xlsx/Credit Memo Batch TEMPLATE.xlsx`. The allowed list is `app/template_rules.py`. Build exports by copying the template, never by inventing columns, sheets or formulas.
- Store the source filename for every imported row.
- Never delete or overwrite uploaded business data (`data/`) without explicit approval.
- Validate uploads before writing records to SQLite.
- One importer module per raw file type in `app/importers/`; register it in `app/importers/__init__.py`.
- No long texts on screen: explanations go in `{{ info('text') }}` from `_macros.html`.
- Keep the UI understandable for non-technical accounting users.

## Verification

Run `python -m py_compile app/*.py app/importers/*.py` after Python changes, and `python -m pytest tests` when tests exist.
Recheck protocols (`app/recheck.py`, rules in `PROTOCOLS.md`; update `PROTOCOLS.md` whenever a protocol is added or changed): every protocol has a test in `tests/test_recheck.py` (a clean batch must give zero errors; each mistake must be caught by its protocol). When a protocol is added or changed, add or update its test and run `python -m pytest tests` before finishing.
Run `start_dashboard.bat` (or `python -m uvicorn app.main:app --reload`) to launch on the host/port in `config.py`.

## Tables and viewer

- Add a table: add it to `app/importers/specs.py` (`TABLES`, `KIND_TABLES`). The SQLite schema and the viewer columns both come from there.
- Every stored row keeps `source_file_id` and `source_file`; re-importing a file replaces only that file's rows.
- The A/R PDF is parsed in the browser with pdf.js and `static/parse_report.js` (unchanged from the Report2Excel v7 tool); do not re-implement it in Python.
- `static/viewer_engine.js` and `viewer.css` come from the v7 tool; engine styles are scoped under `.dv`.
- Table downloads use DataTables Buttons 3.1.2 (copy, CSV, XLSX via JSZip 3.10.1, PDF via pdfmake 0.2.7, print, column visibility) from `static/vendor/`, plus a custom XLS (Excel 2003 XML) button in `viewer.js`. Exports use the rows as shown (filtered and sorted). PDF/print of more than 3,000 rows asks first.
- Double-clicking a table cell copies its value (`viewer.js`).
- Data loading pipeline: the viewer first fetches `/api/files/{id}/data?limit=100` (instant), then streams the full data from a prebuilt gzip cache in `data/cache/` (built after each import by `app/data_cache.py`, ETag/304 on repeat views, cleared when rows change or the file is deleted). Any load or build error shows a message and a Try again button; the viewer never stays on "Loading data…".
