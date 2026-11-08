"""Step 1 of the pipeline: read the inward sheet (CSV or Excel) into clean records.

Every later stage only sees the clean "record" dicts produced here, so the rest of
the program never has to care whether the data came from .csv or .xlsx.
"""
import csv
import datetime as dt
import io
import re

import openpyxl

# Column names in the lab's inward sheet -> our internal names.
# To support a new sheet layout, add its header spelling here. Nothing else changes.
ALIASES = {
    "sr_no": ["sr no", "sr no.", "s no", "s. no.", "sr"],
    "rec_date": ["rec. date", "rec date", "received date", "date of sample"],
    "ar_no": ["a.r. no.", "a.r. no", "a.r no.", "a.r no", "ar no", "ar no."],
    "product_name": ["product name", "product", "product/sample name"],
    "sku": ["sku", "pack size", "pack"],
    "batch_no": ["batch no", "batch no.", "batch", "batch number"],
    "mfg": ["mfg", "mfg date", "mfg."],
    "exp": ["exp", "exp date", "expiry", "expiry date"],
    "tvc": ["tvc", "tbc", "total bacterial count"],
    "ym": ["y&m", "ym", "y & m", "tyc", "total yeast & mould count"],
    "pathogens": ["pathogens", "pathogen"],
    "analysis_date": ["analysis date", "date of analysis"],
    "release_date": ["release date", "date of completion", "completion date"],
    "done_by": ["done by", "analysed by"],
}
_LOOKUP = {alias: key for key, names in ALIASES.items() for alias in names}

DATE_KEYS = ("rec_date", "mfg", "exp", "analysis_date", "release_date")
RESULT_KEYS = ("tvc", "ym", "pathogens")


def _norm_header(value):
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def _clean(value):
    """Turn any cell value into a tidy string ('' for empty)."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)  # 2114980.0 -> 2114980
    return re.sub(r"\s+", " ", str(value)).strip()


def _to_date(value):
    """Return an ISO date string 'YYYY-MM-DD' if value is a date, else '' or the raw text.

    Excel gives real datetimes; CSV gives text like 23/09/2026 (day first, Indian style).
    Text we cannot understand (e.g. '09/26' = month/year) is kept as-is.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    text = _clean(value)
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return text


def _read_rows(filename, data):
    """Return a list of rows (each a list of raw cell values) from csv/xlsx bytes."""
    name = filename.lower()
    if name.endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
        ws = wb.worksheets[0]
        return [list(r) for r in ws.iter_rows(values_only=True)]
    if name.endswith(".csv"):
        for enc in ("utf-8-sig", "cp1252"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        return list(csv.reader(io.StringIO(text)))
    raise ValueError("Please upload a .csv or .xlsx file.")


def _find_header(rows):
    """The sheet has title lines above the real header, so look for the header row."""
    for i, row in enumerate(rows[:30]):
        mapped = {_LOOKUP.get(_norm_header(c)) for c in row}
        if {"ar_no", "batch_no"} <= mapped:
            return i, {j: _LOOKUP[_norm_header(c)] for j, c in enumerate(row) if _norm_header(c) in _LOOKUP}
    raise ValueError(
        "Could not find the header row. The file needs at least 'A.R. No.' and 'Batch No' columns."
    )


def load_records(filename, data):
    """Parse an inward sheet. Returns (records, info)."""
    rows = _read_rows(filename, data)
    header_idx, colmap = _find_header(rows)
    records, skipped = [], 0
    for row in rows[header_idx + 1:]:
        rec = {key: "" for key in ALIASES}
        for j, key in colmap.items():
            if j < len(row):
                rec[key] = _to_date(row[j]) if key in DATE_KEYS else _clean(row[j])
        # Skip junk rows such as "Month of October 2026" or blank spacer rows.
        if not rec["ar_no"] or not rec["batch_no"]:
            skipped += 1
            continue
        rec["ready"] = all(rec[k] for k in RESULT_KEYS) and bool(rec["analysis_date"] and rec["release_date"])
        rec["id"] = len(records)
        records.append(rec)
    info = {
        "filename": filename,
        "total": len(records),
        "ready": sum(r["ready"] for r in records),
        "skipped": skipped,
    }
    return records, info


def fmt_date(iso):
    """'2026-07-30' -> '30/07/2026'. Anything else is returned unchanged."""
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else (iso or "")
