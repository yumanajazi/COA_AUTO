"""Step 3: pour the resolved COA data into the company's own Word template.

We never rebuild the layout. We open your real COA (template_docx/COA_TEMPLATE.docx),
change only the text inside it, and save a copy. That is why the output keeps your
logo, borders, fonts and header exactly as they are.
"""
import copy
import os
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

TEMPLATE = Path(__file__).resolve().parent.parent / "template_docx" / "COA_TEMPLATE.docx"

# label printed in the template header -> key in resolved["header"]
HEADER_LABELS = {
    "Product/Sample Name": "product_name",
    "A.R No.": "ar_no",
    "Batch No.": "batch_no",
    "Mfg Date": "mfg",
    "Pack Size": "pack_size",
    "Expiry Date": "exp",
    "Date of Sample": "date_sample",
    "Sample Qty.": "sample_qty",
    "Date of Analysis": "date_analysis",
    "Date of Completion": "date_completion",
    "Packaging Details": "packaging",
    "Submitted To": "submitted_to",
    "Sample Reg. Sr. No.": "sample_reg",
}


def _set_text(t, text):
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _fill_header(header_el, values):
    """Each value sits in the first non-blank text run after its label."""
    texts = list(header_el.iter(qn("w:t")))
    for label, key in HEADER_LABELS.items():
        for i, t in enumerate(texts):
            if (t.text or "").strip() == label:
                for nxt in texts[i + 1:]:
                    if (nxt.text or "").strip():
                        _set_text(nxt, values[key])
                        break
                break
        else:
            raise KeyError(f"Label '{label}' not found in template header")


def _set_cell(cell, text):
    """Replace the visible text of a table cell but keep its formatting."""
    first = True
    for t in cell._tc.iter(qn("w:t")):
        if first and (t.text or "").strip():
            _set_text(t, text)
            first = False
        elif (t.text or "").strip() or not first:
            t.text = ""
    if first:  # no text run at all: add one to the last paragraph
        cell.paragraphs[-1].add_run(text)


def _fill_table(doc, rows):
    table = doc.tables[0]
    tr = [r._tr for r in table.rows]
    # Template rows: 0 header | 1 TBC | 2 Y&M | 3 "Pathogens" | 4..8 pathogen rows
    templates = {"param": copy.deepcopy(tr[4]), "section": copy.deepcopy(tr[3])}
    for el in tr[1:]:
        table._tbl.remove(el)
    for row in rows:
        new = copy.deepcopy(templates[row["type"]])
        table._tbl.append(new)
    # python-docx rows are now the freshly appended copies
    for row, tbl_row in zip(rows, table.rows[1:]):
        if row["type"] == "param":
            c = tbl_row.cells
            _set_cell(c[0], row["sno"])
            _set_cell(c[1], row["label"])
            _set_cell(c[2], row["spec"])
            _set_cell(c[3], row["obs"])
        else:
            _set_cell(tbl_row.cells[0], row["label"])


def _fill_remark(doc, remark):
    if not remark:
        return
    for t in doc.element.body.iter(qn("w:t")):
        if t.text and "Complies/Does not Complies" in t.text:
            _set_text(t, t.text.replace("Complies/Does not Complies", remark))
            return


def _fill_footer(footer_el, sig):
    prefixes = {"ANALYSED BY:": sig["analysed"], "REVIEWED BY:": sig["reviewed"], "APPROVED BY:": sig["approved"]}
    for t in footer_el.iter(qn("w:t")):
        for prefix, name in prefixes.items():
            if (t.text or "").strip().startswith(prefix):
                _set_text(t, f"{prefix}  {name.upper()}")


def build_docx(resolved, out_path):
    doc = Document(str(TEMPLATE))
    section = doc.sections[0]
    _fill_header(section.header._element, resolved["header"])
    _fill_table(doc, resolved["rows"])
    _fill_remark(doc, resolved["remark"])
    _fill_footer(section.footer._element, resolved["signatories"])
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    return out_path
