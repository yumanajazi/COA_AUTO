# Learning guide: how COA-Automation works, and how to grow it

## 1. The one idea behind the whole project

Every automation like this is the same four steps:

```
 INPUT  ->  CLEAN  ->  DECIDE  ->  OUTPUT
 inward     loader.py   resolver.py   docx_builder.py (+ converter.py)
 sheet      "what rows   "what text    "put it in the Word
            exist?"       goes where?"   template"
```

Why this split matters: each step can change without breaking the others.
Tomorrow's input might be a lab machine's export instead of Excel. Only `loader.py` changes.
Tomorrow's COA might have a new layout. Only the template changes.

## 2. The most important design choice: ONE source of truth

`resolve(record, settings)` returns the final text of the COA as a plain dictionary.
Both the **live preview** (`_preview.html`) and the **Word file** (`docx_builder.py`) are drawn from
that same dictionary. So the preview can never disagree with the download.
If you ever add a field, you add it in `resolver.py` and then show it in the two places. That's it.

## 3. File-by-file, in the order to read them

| Read | File | Concept you learn |
|---|---|---|
| 1 | `coa/loader.py` | Reading Excel/CSV, aliases for column names, cleaning messy data (junk rows, `2114980.0`, day-first dates) |
| 2 | `coa/resolver.py` | Plain dicts as data, settings with defaults, rules (compliance check: `NMT 100` vs `<10`) |
| 3 | `coa/docx_builder.py` | Editing a real Word file: Word is XML; we change text inside runs and clone table rows |
| 4 | `app.py` | Flask routes, sessions, GET vs POST, returning JSON |
| 5 | `templates/*.html` | Jinja templates: loops, macros, filters |
| 6 | `static/js/customize.js` | The live preview loop: read form -> `fetch` -> replace HTML (debounced) |
| 7 | `tests/test_pipeline.py` | Why tests: change the code boldly, run one command, know nothing broke |

## 4. Exercises (each is 10 to 30 minutes and teaches one thing)

1. **Add a new field.** Add "Done by" to the COA header. (Add to `FIELDS` in `resolver.py`, add the label to
   `HEADER_LABELS` in `docx_builder.py`, add a row in `_preview.html`.) You'll touch every layer once.
2. **Change a rule.** Make Y&M limit `NMT 100`. Watch the remark flip. Then write a test for it.
3. **Add a second COA template** (e.g. for water or raw material). Copy the template, add a dropdown on page 2, pass the
   template path into `build_docx`.
4. **Edit the template in Word.** Change the logo or font in `COA_TEMPLATE.docx`. Run again. Nothing in code changes.
5. **Break it on purpose.** Upload a CSV with a missing column. Read the error, then improve the message in `loader.py`.

## 5. Things the prototype assumes (please confirm, each is one line to change)

- `MU-I/07-26/0122` becomes `RSH-W/QC-MUI/07-26/0122` (copied from your sample COA; prefix editable on page 3).
- **Sample Qty** is not in the inward sheet, so it defaults to `NA`. Type it on page 3 or add a column.
- **C.albicans** is `NA` (as on your sample COA); the other four pathogens copy the single "Pathogens" column.
- Pack Size = the `SKU` column. Sample Reg. Sr. No. = `SR NO`. Date of Completion = `Release Date`.
- Remark is judged automatically (Complies / Does not Complies) from the results against the limits; you can force it.
- Signature names (Raksha Singh, Saransh, Asheesh Kumar) come from the template and are editable.
- 141 rows in your inward sheet have no results yet. They are shown as "no results yet" and hidden by default.
- File names follow your sample: `04_BADFF10026` = Sr No + Batch.

## 6. Roadmap: from prototype to "end to end, automatic"

**Phase 1: Daily-use tool (1 to 2 weeks)**
- Run on one office PC; double-click launcher (`python app.py` in a `.bat` file).
- Add a login (Flask-Login) and a "who generated this" log.
- Keep a history table (SQLite) of every COA generated: AR No., batch, who, when. Prevents duplicates.

**Phase 2: Remove the upload step (2 to 4 weeks)**
- Read the inward workbook straight from a shared folder or Google Sheet (gspread) instead of uploading.
- "Watch folder": when a new row gets results + release date, a COA is generated automatically.
- Replace "type your own value" with per-product defaults (a small products table: product, pack, spec limits).

**Phase 3: Quality and trust**
- Validation rules before release: result present, dates in order (analysis <= release), batch format, duplicate A.R. No.
- Review / approve workflow: Analyst generates, Reviewer approves, and the approval is stamped on the COA.
- Real PDF with digital signature or QR code that opens a verification page.

**Phase 4: Machine / LIMS integration**
- Machines usually export CSV, XML or a database table. Add a new loader that maps their columns to the same record dict.
  Nothing after the loader changes. This is why step 1 exists.
- Move from Flask's dev server to a proper server (gunicorn / waitress) and Docker; add a database (PostgreSQL).
- Email the finished COA to the requester automatically.

## 7. Common problems

- **PDF button is greyed out**: install LibreOffice. On Windows the app finds it automatically.
- **Preview looks slightly different from Word**: the preview is a drawing of the template. The Word/PDF file is the truth.
  If you restyle the template a lot, mirror it in `app.css` (`.coa` section).
- **A header label is "not found in template"**: you renamed a label in the Word template. Update `HEADER_LABELS` to match.
- **Sessions**: the prototype keeps each user's uploads in `instance/sessions/`. Safe to delete anytime.

## 8. Questions that will shape the next version

1. Is one COA template enough, or are there several (water, raw material, finished goods, different units)?
2. Where do results originate: typed by an analyst into the sheet, or exported by an instrument? Which one first?
3. Who approves COAs today, and should the app record that?
4. Should specifications (limits) depend on the product (e.g. sunscreen vs powder)?
5. How many COAs per day or month, and who will use the tool?
