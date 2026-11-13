"""Source-file registry. One entry per raw file type; each gets its own importer module later.

To add a new source: add an entry here, then add app/importers/<key>.py with an import function.
"""

SOURCES = [
    {"key": "ageing", "label": "Ageing report", "role": "raw", "accept": ".xlsx,.xls,.xlsm,.xlsb",
     "info": "Full Ageing, Account Analyst List and the summary tabs (Ageing xlsx)."},
    {"key": "orders", "label": "Sale Report", "role": "raw", "accept": ".csv,.xlsx,.xls",
     "info": "Open orders report (LW9301 csv): SKU, qty, price, store, order and invoice numbers."},
    {"key": "ar_pdf", "label": "Report 2", "role": "raw", "accept": ".pdf,.xlsx",
     "info": "Reports 2 PDF, or the Excel export from the Report2Excel tool in tools/report2_pdf_importer."},
    {"key": "reasons", "label": "Reason codes", "role": "raw", "accept": ".csv,.xlsx,.xls",
     "info": "The reason code list (code and description). Import a new file to update it, or use Edit codes to change it by hand."},
    {"key": "validate", "label": "File to validate", "role": "validate", "accept": ".xlsx",
     "info": "A filled credit memo batch to recheck against the imported data."},
]
SOURCE_BY_KEY = {s["key"]: s for s in SOURCES}
