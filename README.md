# Credit Department

FastAPI app for building and rechecking credit memo batches in the exact Excel template.

## Run

Double-click `start_dashboard.bat`, then open http://192.168.2.123:8020. Or manually:

```powershell
python -m pip install -r requirements.txt
python -m uvicorn app.main:app --host 192.168.2.123 --port 8020 --reload
```

## Pages

- **Dashboard** — files per source type and the workflow.
- **Import center** — upload the five source files (saved in `data/uploads/`).
- **Create batch** — demo of the template form; totals use only the template formulas.
- **Recheck** — planned validation checks (not built yet).
- **Template rules** — the sheets, columns and formulas that exports may use.

See `STRUCTURE_BLUEPRINT_OUTLINE.txt` for the folder layout and `AGENTS.md` for the working rules.
