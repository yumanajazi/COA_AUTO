"""Run with:  python -m pytest tests -q     (or just: python tests/test_pipeline.py)"""
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from coa.docx_builder import build_docx
from coa.loader import load_records
from coa.resolver import default_settings, pack_unit, resolve

SAMPLE = Path(__file__).resolve().parent.parent / "sample_data" / "INWARD_sample.xlsx"


def _records():
    return load_records(SAMPLE.name, SAMPLE.read_bytes())[0]


def test_loader_skips_junk_rows_and_counts():
    recs, info = load_records(SAMPLE.name, SAMPLE.read_bytes())
    assert info["total"] == 1100 and info["ready"] == 959
    assert all(r["ar_no"] and r["batch_no"] for r in recs)


def test_csv_with_day_first_dates():
    csv = (b"A.R. No.,Product Name,Batch No,TVC,Y&M,Pathogens,Rec. Date,Analysis Date,Release Date\n"
           b"MU-I/09-26/0001,Test cream,B1,<10,<10,Absent,23/09/2026,24/09/2026,30/09/2026\n")
    recs, _ = load_records("t.csv", csv)
    assert recs[0]["rec_date"] == "2026-09-23" and recs[0]["ready"]


def test_resolve_matches_sample_coa_conventions():
    s = default_settings(); s["remarks"] = "auto"
    res = resolve(_records()[0], s)
    h = res["header"]
    assert h["ar_no"] == "RSH-W/QC-MUI/07-26/0122"      # MU-I -> RSH-W/QC-MUI
    assert h["date_sample"] == "30/07/2026" and h["mfg"] == "NA"
    obs = {r["label"]: r["obs"] for r in res["rows"] if r["type"] == "param"}
    assert obs["Total Bacterial Count"] == "<10cfu/" + pack_unit(_records()[0]["sku"])
    assert obs["C.albicans"] == "NA" and obs["E. coli"] == "Absent"
    assert res["remark"] == "Complies"


def test_failing_result_flips_remark():
    s = default_settings(); s["remarks"] = "auto"
    s["params"]["tbc"].update(use_file=False, manual="250cfu/g")
    assert resolve(_records()[0], s)["remark"] == "Does not Complies"


def test_docx_contains_new_values(tmp_path):
    s = default_settings()
    s["fields"]["batch_no"].update(use_file=False, manual="ZZ999")
    s["params"]["salm"]["include"] = False
    s["others"] = [{"label": "Klebsiella", "spec": "Should be absent", "obs": "Absent"}]
    out = build_docx(resolve(_records()[0], s), str(tmp_path / "x.docx"))
    with zipfile.ZipFile(out) as z:
        body = z.read("word/document.xml").decode()
        header = z.read("word/header1.xml").decode()
    assert "ZZ999" in header and "Klebsiella" in body and "Salmonella" not in body


def test_unit_follows_pack_size_with_manual_override():
    assert pack_unit("100 ml") == "ml" and pack_unit("50 gm") == "g" and pack_unit("15ml") == "ml"
    assert pack_unit("NA") is None
    rec = dict(_records()[0], sku="10g", tvc="<10")
    obs = lambda s: next(r["obs"] for r in resolve(rec, s)["rows"] if r["type"] == "param")
    assert obs(default_settings()) == "<10cfu/g"
    rec["sku"] = "10 ml"
    assert obs(default_settings()) == "<10cfu/ml"
    s = default_settings(); s["unit"] = "g"
    assert obs(s) == "<10cfu/g"


def test_category_qty_and_packaging_defaults():
    rec = _records()[0]
    for cat, qty in (("FG", "10g"), ("RM", "10g"), ("WS", "1ml")):
        h = resolve(rec, default_settings(), cat)["header"]
        assert h["sample_qty"] == qty and h["packaging"] == "Original Packaging"
    s = default_settings(); s["fields"]["packaging"]["manual"] = "Sealed bottle"
    assert resolve(rec, s, "FG")["header"]["packaging"] == "Sealed bottle"


def test_category_detection_on_names_and_batches():
    from coa.loader import detect_category
    assert detect_category({"product_name": "DM Water (S.P. No.-03)", "batch_no": "Aug/F1-01"}) == "WS"
    assert detect_category({"product_name": "Dot & Key Watermelon Cooling", "batch_no": "BAOYF10092"}) == "FG"
    assert detect_category({"product_name": "Joy glow", "batch_no": "RSB 26-07 (BABJF10445)"}) == "FG"
    assert detect_category({"product_name": "Aloevera Powder", "batch_no": "A2JJZ032"}) == "RM"


def test_header_label_split_over_runs_and_remark_line(tmp_path):
    out = build_docx(resolve(_records()[0], default_settings()), str(tmp_path / "y.docx"))
    from docx import Document
    d = Document(out)
    assert any(p.text.strip().endswith("2. Above test done by as per In-house STP.") for p in d.paragraphs)


if __name__ == "__main__":
    import tempfile
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn(Path(tempfile.mkdtemp())) if "tmp_path" in fn.__code__.co_varnames[:1] else fn()
            print("ok ", name)
