"""Step 2: turn (record + user choices) into the exact text that goes on the COA.

This is the heart of the project. Both the live HTML preview and the Word/PDF file are
built from the dict returned by `resolve()`, so what you see is what you download.
"""
import copy
import re

from .loader import fmt_date

# ---- Header fields -----------------------------------------------------------------
# key, label on the COA, source column in the record (None = not in the inward sheet),
# group shown in the settings panel, default manual text
FIELDS = [
    ("product_name",    "Product Name",          "product_name",  "Product information", ""),
    ("batch_no",        "Batch No.",             "batch_no",      "Product information", ""),
    ("pack_size",       "Pack Size",             "sku",           "Product information", ""),
    ("ar_no",           "A.R. No.",              "ar_no",         "Product information", ""),
    ("sample_reg",      "Sample Reg. Sr. No.",   "sr_no",         "Product information", ""),
    ("mfg",             "Mfg Date",              "mfg",           "Product information", ""),
    ("exp",             "Expiry Date",           "exp",           "Product information", ""),
    ("packaging",       "Packaging Details",     None,            "Product information", ""),
    ("submitted_to",    "Submitted To",          None,            "Product information", "RSH Global Unit-I"),
    ("sample_qty",      "Sample Qty.",           None,            "Testing details",     ""),
    ("date_sample",     "Date of Sample",        "rec_date",      "Dates",               ""),
    ("date_analysis",   "Date of Analysis",      "analysis_date", "Dates",               ""),
    ("date_completion", "Date of Completion",    "release_date",  "Dates",               ""),
]
# Boxes left empty on page 3 use these automatic values. Type something to override.
DEFAULT_PACKAGING = "Original Packaging"
FIELD_HINTS = {
    "packaging": "Automatic: Original Packaging (type to change)",
    "sample_qty": "Automatic by category: FG 10g, RM 10g, WS 1ml (type to change)",
}

# Sample categories. 'qty' is the automatic Sample Qty. for that category.
# Parameters are shared by all categories; use the "RM & WS" tick on page 3 for one-off changes.
CATEGORIES = {"FG": "Finished Goods", "RM": "Raw Material", "WS": "Water Sample"}
CATEGORY_PRESETS = {
    "FG": {"qty": "10g"},
    "RM": {"qty": "10g"},
    "WS": {"qty": "1ml"},
}
DATE_FIELDS = {"mfg", "exp", "date_sample", "date_analysis", "date_completion"}

# ---- Test parameters ---------------------------------------------------------------
# key, label, specification text, source column, default "take result from file?", default manual text
PARAMS = [
    ("tbc",   "Total Bacterial Count",    "NMT 100cfu/g or ml", "tvc",       True,  ""),
    ("ym",    "Total Yeast & Mould Count", "NMT 10cfu/g or ml",  "ym",        True,  ""),
    ("ecoli", "E. coli",                  "Should be absent",   "pathogens", True,  ""),
    ("paeru", "P. aeruginosa",            "Should be absent",   "pathogens", True,  ""),
    ("saur",  "S. aureus",                "Should be absent",   "pathogens", True,  ""),
    ("calb",  "C.albicans",               "Should be absent",   "pathogens", False, "NA"),
    ("salm",  "Salmonella sp.",           "Should be absent",   "pathogens", True,  ""),
]
PATHOGEN_KEYS = {"ecoli", "paeru", "saur", "calb", "salm"}
AR_PREFIX = "RSH-W/QC-"


def default_settings():
    return {
        "fields": {k: {"use_file": src is not None and not default, "manual": default}
                   for k, _, src, _, default in FIELDS},
        "params": {k: {"include": True, "spec": spec, "use_file": use_file, "manual": manual}
                   for k, _, spec, _, use_file, manual in PARAMS},
        "others": [],           # extra rows typed by hand: [{"label","spec","obs"}]
        "unit": "auto",         # auto (from pack size) | g | ml
        "remarks": "keep",      # auto | Complies | Does not Complies | keep
        "ar_prefix": AR_PREFIX,
        "signatories": {"analysed": "RAKSHA SINGH", "reviewed": "SARANSH", "approved": "ASHEESH KUMAR"},
    }


def merge_settings(incoming):
    """Overlay the browser's settings onto the defaults so missing keys never crash us."""
    base = default_settings()
    if not isinstance(incoming, dict):
        return base
    for section in ("fields", "params"):
        for key, val in (incoming.get(section) or {}).items():
            if key in base[section] and isinstance(val, dict):
                base[section][key].update(val)
    base["others"] = [o for o in (incoming.get("others") or []) if isinstance(o, dict)]
    for key in ("remarks", "ar_prefix", "unit"):
        if isinstance(incoming.get(key), str):
            base[key] = incoming[key]
    base["signatories"].update({k: v for k, v in (incoming.get("signatories") or {}).items()
                                if k in base["signatories"] and isinstance(v, str)})
    return base


# ---- Value formatting --------------------------------------------------------------
def pack_unit(sku):
    """'100 ml' -> 'ml', '50 gm' -> 'g', '15ml' -> 'ml'. None when the pack size has no unit."""
    m = re.search(r"\d\s*(ml|kg|gm|mg|g|l)\b", (sku or "").lower())
    if not m:
        return None
    return "ml" if m.group(1) in ("ml", "l") else "g"


def _fmt_result(raw, unit="g"):
    """'<10' -> '<10cfu/g' (or cfu/ml). Text such as 'Absent' is left alone."""
    raw = (raw or "").strip()
    return raw + f"cfu/{unit}" if re.fullmatch(r"[<>]?\s*\d+(\.\d+)?", raw) else raw


def _ar_number(raw, prefix):
    """Inward 'MU-I/07-26/0122' -> COA 'RSH-W/QC-MUI/07-26/0122'."""
    return prefix + raw.replace("MU-I", "MUI") if raw else ""


def _field_value(key, record, cfg, prefix):
    src = next(s for k, _, s, _, _ in FIELDS if k == key)
    if src and cfg["use_file"]:
        raw = record.get(src, "")
        if key in DATE_FIELDS:
            raw = fmt_date(raw)
        elif key == "ar_no":
            raw = _ar_number(raw, prefix)
        return raw
    return cfg["manual"].strip()


def _complies(spec, obs):
    """True / False / None (None = cannot judge, e.g. 'NA')."""
    o, s = obs.strip().lower(), spec.lower()
    if not o or o == "na":
        return None
    if "absent" in s:
        return {"absent": True, "present": False}.get(o)
    limit = re.search(r"nmt\s*(\d+(?:\.\d+)?)", s)
    num = re.fullmatch(r"([<>]?)\s*(\d+(?:\.\d+)?)\s*(?:cfu.*)?", o)
    if limit and num:
        sign, value = num.group(1), float(num.group(2))
        if sign == ">":
            return False
        return value <= float(limit.group(1))
    return None


def resolve(record, settings, category=None):
    """Return everything the COA needs, already formatted as final text.

    category: 'FG' / 'RM' / 'WS'. If None, the category detected when the file was read is used.
    """
    s = merge_settings(settings)
    warnings = []
    category = category if category in CATEGORIES else record.get("category", "FG")
    preset = CATEGORY_PRESETS[category]

    # unit of the observations: from the pack size, unless the user chose g or ml by hand
    if s["unit"] in ("g", "ml"):
        unit = s["unit"]
    else:
        unit = pack_unit(record.get("sku", ""))
        if unit is None:
            unit = "ml" if category == "WS" else "g"
            warnings.append(f"Pack size '{record.get('sku') or 'empty'}' has no g/ml: used cfu/{unit}")

    header = {}
    for key, label, src, _, _ in FIELDS:
        value = _field_value(key, record, s["fields"][key], s["ar_prefix"])
        if not value and key == "sample_qty":
            value = preset["qty"]
        elif not value and key == "packaging":
            value = DEFAULT_PACKAGING
        if not value:
            value = "NA"
            if key in ("product_name", "batch_no", "ar_no", "date_sample", "date_analysis", "date_completion"):
                warnings.append(f"{label} is empty in the uploaded file")
        header[key] = value

    rows, verdicts, number, pathogen_heading_added = [], [], 0, False

    def add(label, spec, obs, pathogen):
        nonlocal number, pathogen_heading_added
        if pathogen and not pathogen_heading_added:
            rows.append({"type": "section", "label": "Pathogens"})
            pathogen_heading_added = True
        number += 1
        rows.append({"type": "param", "sno": f"{number:02d}", "label": label, "spec": spec, "obs": obs})
        verdicts.append(_complies(spec, obs))

    for key, label, _, src, _, _ in PARAMS:
        cfg = s["params"][key]
        if not cfg["include"]:
            continue
        if cfg["use_file"]:
            obs = _fmt_result(record.get(src, ""), unit)
            if not obs:
                obs = "NA"
                warnings.append(f"No result for {label} in the uploaded file")
        else:
            obs = cfg["manual"].strip() or "NA"
        add(label, cfg["spec"].strip(), obs, key in PATHOGEN_KEYS)
    for other in s["others"]:
        if (other.get("label") or "").strip():
            add(other["label"].strip(), (other.get("spec") or "").strip(),
                (other.get("obs") or "").strip() or "NA", True)

    mode = s["remarks"]
    if mode == "auto":
        judged = [v for v in verdicts if v is not None]
        unknown = any(v is None for v in verdicts) and not judged
        if unknown or not judged:
            remark = None
        else:
            remark = "Complies" if all(judged) else "Does not Complies"
    elif mode == "keep":
        remark = None
    else:
        remark = mode

    return {
        "header": header,
        "rows": rows,
        "remark": remark,              # None -> leave template wording "Complies/Does not Complies"
        "signatories": s["signatories"],
        "warnings": warnings,
        "file_stem": _file_stem(record),
        "category": category,
        "unit": unit,
    }


def _file_stem(record):
    sr = record.get("sr_no", "")
    sr = f"{int(sr):02d}" if sr.isdigit() else sr
    stem = f"{sr}_{record.get('batch_no', 'COA')}"
    return re.sub(r"[^A-Za-z0-9_.-]+", "-", stem).strip("-") or "COA"
