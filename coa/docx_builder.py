"""Step 3: pour the resolved COA data into the company's own Word template.

We never rebuild the layout. We open your real COA (template_docx/COA_TEMPLATE.docx),
change only the text inside it, and save a copy. That is why the output keeps your
logo, borders, fonts and header exactly as they are.
"""
import copy
import os
from pathlib import Path

from docx import Document
from lxml import etree
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


_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_SIDE_PARENTS = {"tblCellMar", "tcMar", "tcBorders", "tblBorders", "pBdr"}


def _word_compatible(root):
    """Make the XML acceptable to every Word version."""
    for tag, new in (("start", "left"), ("end", "right")):
        for el in list(root.iter(f"{{{_W}}}{tag}")):
            parent = el.getparent()
            if parent is not None and etree.QName(parent).localname in _SIDE_PARENTS:
                el.tag = f"{{{_W}}}{new}"
    for el in root.iter("{*}docPr", "{*}cNvPr"):
        if "title" in el.attrib:
            del el.attrib["title"]

HEADER_FILL = "BDD7EE"   # colour of the S.No./Parameters/Specification/Observations row (hex, no #)
_TCPR_AFTER_SHD = ("noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark")
_TCPR_AFTER_VALIGN = ("hideMark",)


def _insert_in_tcpr(tcpr, new, before_names):
    """Put <new> before the first child whose name is in before_names (Word wants a fixed order)."""
    for child in tcpr:
        if etree.QName(child).localname in before_names:
            child.addprevious(new)
            return
    tcpr.append(new)


def _style_table(table):
    """Vertically centre every cell of the results table and colour its header row."""
    for r_index, tr in enumerate(table._tbl.tr_lst):
        for tc in tr.tc_lst:
            tcpr = tc.get_or_add_tcPr()
            valign = tcpr.find(f"{{{_W}}}vAlign")
            if valign is None:
                valign = etree.SubElement(tcpr, f"{{{_W}}}vAlign")
                tcpr.remove(valign)
                _insert_in_tcpr(tcpr, valign, _TCPR_AFTER_VALIGN)
            valign.set(f"{{{_W}}}val", "center")
            if r_index == 0:                                  # header row only
                shd = tcpr.find(f"{{{_W}}}shd")
                if shd is None:
                    shd = etree.SubElement(tcpr, f"{{{_W}}}shd")
                    tcpr.remove(shd)
                    _insert_in_tcpr(tcpr, shd, _TCPR_AFTER_SHD)
                shd.set(f"{{{_W}}}val", "clear")
                shd.set(f"{{{_W}}}color", "auto")
                shd.set(f"{{{_W}}}fill", HEADER_FILL)


def _right_align_doc_numbers(root):
    """Put the form numbers (RSH-W/QA/...) at the right margin."""
    for p in root.iter(f"{{{_W}}}p"):
        texts = list(p.iter(f"{{{_W}}}t"))
        if not any("RSH-W/QA/" in (t.text or "") for t in texts):
            continue
        for r in list(p.iter(f"{{{_W}}}r")):                  # remove runs that hold only spaces
            if all(not (t.text or "").strip() for t in r.iter(f"{{{_W}}}t")) and r.find(f"{{{_W}}}t") is not None:
                r.getparent().remove(r)
        ppr = p.find(f"{{{_W}}}pPr")
        if ppr is None:
            ppr = etree.Element(f"{{{_W}}}pPr")
            p.insert(0, ppr)
        for old in ppr.findall(f"{{{_W}}}jc"):
            ppr.remove(old)
        jc = etree.Element(f"{{{_W}}}jc")
        jc.set(f"{{{_W}}}val", "right")
        for child in ppr:
            if etree.QName(child).localname in ("textDirection", "textAlignment", "textboxTightWrap",
                                                "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"):
                child.addprevious(jc)
                break
        else:
            ppr.append(jc)


def build_docx(resolved, out_path):
    doc = Document(str(TEMPLATE))
    section = doc.sections[0]
    _fill_header(section.header._element, resolved["header"])
    _fill_table(doc, resolved["rows"])
    _fill_remark(doc, resolved["remark"])
    _fill_footer(section.footer._element, resolved["signatories"])
    _style_table(doc.tables[0])
    _right_align_doc_numbers(section.header._element)
    _right_align_doc_numbers(section.footer._element)
    for el in (doc.element, section.header._element, section.footer._element):
        _word_compatible(el)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    return out_path
