# COA-Automation (prototype)

Upload the lab's inward sheet, pick samples, choose what to change, watch a live preview,
and download finished Certificates of Analysis (Word or PDF) in your own company format.

## Run it

```bash
pip install -r requirements.txt
python app.py
# open http://127.0.0.1:5000  ->  click "try it with the sample inward file"
```

PDF output needs LibreOffice installed (free). Without it the PDF option is greyed out and Word still works.

## The four pages (from your pencil blueprint)

| Page | URL | What it does |
|---|---|---|
| 1 Upload | `/` | Upload `.xlsx` or `.csv` (or use the sample) |
| 2 Review | `/select` | See every row, search/filter, tick the ones to convert, pick Word or PDF |
| 3 Customize | `/customize` | Tick = take from file, untick = type your own value. Live preview, page through samples, add extra tests by hand |
| 4 Download | `/download` | Save each COA, download ticked ones, or the whole batch as a ZIP |

## Folder map

```
app.py                  web pages and routes (the "waiter")
coa/loader.py           reads CSV/Excel -> clean records
coa/resolver.py         record + your choices -> final text   <- the brain
coa/docx_builder.py     pours that text into your Word template
coa/converter.py        Word -> PDF via LibreOffice
template_docx/          COA_TEMPLATE.docx  (your real COA; edit it in Word to restyle everything)
templates/, static/     HTML, CSS, JS for the 4 pages and the live preview
sample_data/            your inward sheet, for demos
tests/                  python -m pytest tests -q
```

See `LEARNING_GUIDE.md` for how it works and how to grow it.
