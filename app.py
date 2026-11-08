"""COA Automation - Flask web app.

Run:  python app.py      then open http://127.0.0.1:5000

Page flow (matches the pencil blueprint):
  /          1. Upload the inward sheet
  /select    2. See the uploaded rows, tick the ones to convert, choose Word/PDF
  /customize 3. Choose what to change, watch the live preview, press "Process final COA"
  /download  4. Save each COA, or download the whole batch as a ZIP
"""
import io
import json
import secrets
import shutil
import zipfile
from pathlib import Path

from flask import (Flask, abort, jsonify, redirect, render_template, request,
                   send_from_directory, session, url_for)

from coa.converter import docx_to_pdf, pdf_available
from coa.docx_builder import build_docx
from coa.loader import fmt_date, load_records
from coa.resolver import FIELDS, PARAMS, default_settings, merge_settings, resolve

BASE = Path(__file__).resolve().parent
SESSIONS = BASE / "instance" / "sessions"
SAMPLE = BASE / "sample_data" / "INWARD_sample.xlsx"
MAX_COAS_PER_RUN = 300  # safety cap for the prototype

app = Flask(__name__)
app.config["SECRET_KEY"] = secrets.token_hex(16)  # prototype: new key each start
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
app.jinja_env.filters["ddmmyyyy"] = fmt_date


# ---------------------------------------------------------------- per-user workspace
def workdir():
    sid = session.get("sid")
    if not sid:
        sid = session["sid"] = secrets.token_hex(8)
    path = SESSIONS / sid
    (path / "out").mkdir(parents=True, exist_ok=True)
    return path


def load_state():
    f = workdir() / "state.json"
    return json.loads(f.read_text("utf-8")) if f.exists() else {}


def save_state(state):
    (workdir() / "state.json").write_text(json.dumps(state), "utf-8")


def need(*keys):
    """Return state, or None if an earlier step has not been completed yet."""
    state = load_state()
    return state if all(state.get(k) for k in keys) else None


# ---------------------------------------------------------------- page 1: upload
@app.get("/")
def upload_page():
    return render_template("upload.html", step=1, has_sample=SAMPLE.exists(), state=load_state())


def _ingest(filename, data):
    records, info = load_records(filename, data)
    if not records:
        raise ValueError("No sample rows found in that file.")
    shutil.rmtree(workdir() / "out", ignore_errors=True)
    (workdir() / "out").mkdir(exist_ok=True)
    save_state({"records": records, "info": info, "selected": [], "format": "docx"})


@app.post("/upload")
def upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return render_template("upload.html", step=1, has_sample=SAMPLE.exists(), state={},
                               error="Please choose a file first."), 400
    try:
        _ingest(f.filename, f.read())
    except Exception as e:  # show a friendly message instead of a crash page
        return render_template("upload.html", step=1, has_sample=SAMPLE.exists(), state={},
                               error=str(e)), 400
    return redirect(url_for("select_page"))


@app.post("/use-sample")
def use_sample():
    _ingest(SAMPLE.name, SAMPLE.read_bytes())
    return redirect(url_for("select_page"))


# ---------------------------------------------------------------- page 2: review rows
@app.get("/select")
def select_page():
    state = need("records")
    if not state:
        return redirect(url_for("upload_page"))
    return render_template("select.html", step=2, state=state, pdf_ok=pdf_available())


@app.post("/select")
def select_submit():
    state = need("records")
    if not state:
        return redirect(url_for("upload_page"))
    ids = [int(i) for i in request.form.getlist("ids") if i.isdigit()]
    valid = {r["id"] for r in state["records"]}
    ids = [i for i in ids if i in valid][:MAX_COAS_PER_RUN]
    if not ids:
        return render_template("select.html", step=2, state=state, pdf_ok=pdf_available(),
                               error="Tick at least one row to convert."), 400
    fmt = request.form.get("format", "docx")
    state["selected"] = ids
    state["format"] = "pdf" if fmt == "pdf" and pdf_available() else "docx"
    save_state(state)
    return redirect(url_for("customize_page"))


# ---------------------------------------------------------------- page 3: customise + live preview
@app.get("/customize")
def customize_page():
    state = need("records", "selected")
    if not state:
        return redirect(url_for("select_page"))
    by_id = {r["id"]: r for r in state["records"]}
    items = [{"id": i, "sr": by_id[i]["sr_no"], "batch": by_id[i]["batch_no"],
              "product": by_id[i]["product_name"]} for i in state["selected"]]
    settings = state.get("settings") or default_settings()
    return render_template("customize.html", step=3, state=state, items=items, fields=FIELDS,
                           params=PARAMS, settings=merge_settings(settings),
                           capped=len(state["selected"]) >= MAX_COAS_PER_RUN)


@app.post("/api/preview")
def api_preview():
    state = need("records", "selected")
    if not state:
        abort(400)
    body = request.get_json(force=True, silent=True) or {}
    rid = body.get("id")
    if rid not in state["selected"]:
        abort(404)
    record = state["records"][rid]
    data = resolve(record, body.get("settings"))
    # also send what the *file* says for each field, so the form can show it greyed out
    return jsonify(html=render_template("_preview.html", coa=data), warnings=data["warnings"],
                   source=_source_values(record, body.get("settings")))


def _source_values(record, settings):
    from coa.resolver import _ar_number, DATE_FIELDS
    s = merge_settings(settings)
    out = {}
    for key, _, src, _, _ in FIELDS:
        if src:
            raw = record.get(src, "")
            raw = fmt_date(raw) if key in DATE_FIELDS else raw
            if key == "ar_no":
                raw = _ar_number(raw, s["ar_prefix"])
            out[key] = raw or "(empty in file)"
    return out


@app.post("/api/generate")
def api_generate():
    state = need("records", "selected")
    if not state:
        abort(400)
    settings = merge_settings((request.get_json(force=True, silent=True) or {}).get("settings"))
    out = workdir() / "out"
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(exist_ok=True)
    built = []
    for rid in state["selected"]:
        rec = state["records"][rid]
        res = resolve(rec, settings)
        path = build_docx(res, str(out / f"{res['file_stem']}.docx"))
        built.append({"id": rid, "stem": res["file_stem"], "docx": path, "sr": rec["sr_no"],
                      "batch": rec["batch_no"], "product": rec["product_name"], "warnings": res["warnings"]})
    fmt = state["format"]
    if fmt == "pdf":
        try:
            pdfs = docx_to_pdf([b["docx"] for b in built], out)
        except Exception as e:
            return jsonify(error=f"PDF conversion failed: {e}"), 500
        for b in built:
            b["file"] = Path(pdfs[b["docx"]] or b["docx"]).name
        for b in built:  # keep only the requested format on disk
            if b["file"].endswith(".pdf"):
                Path(b["docx"]).unlink(missing_ok=True)
    else:
        for b in built:
            b["file"] = Path(b["docx"]).name
    outputs = [{k: b[k] for k in ("id", "sr", "batch", "product", "file", "warnings")} for b in built]
    state.update(settings=settings, outputs=outputs)
    save_state(state)
    return jsonify(ok=True, redirect=url_for("download_page"))


# ---------------------------------------------------------------- page 4: download
@app.get("/download")
def download_page():
    state = need("outputs")
    if not state:
        return redirect(url_for("customize_page"))
    return render_template("download.html", step=4, state=state)


@app.get("/file/<name>")
def get_file(name):
    state = need("outputs")
    if not state or name not in {o["file"] for o in state["outputs"]}:
        abort(404)
    return send_from_directory(workdir() / "out", name, as_attachment=True)


@app.post("/batch.zip")
def batch_zip():
    state = need("outputs")
    if not state:
        abort(404)
    allowed = {o["file"] for o in state["outputs"]}
    if request.form.get("mode") == "ticked":
        names = [n for n in request.form.getlist("names") if n in allowed]
        if not names:
            return redirect(url_for("download_page"))
    else:
        names = sorted(allowed)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(workdir() / "out" / n, n)
    buf.seek(0)
    from flask import send_file
    return send_file(buf, mimetype="application/zip", as_attachment=True, download_name="COA_batch.zip")


if __name__ == "__main__":
    SESSIONS.mkdir(parents=True, exist_ok=True)
    app.run(debug=True, port=5000)
