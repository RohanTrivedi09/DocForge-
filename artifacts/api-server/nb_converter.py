"""
.ipynb → .docx conversion for WordForge.
Handles markdown, code cells, outputs, images, DataFrames, and TOC.
"""

import io
import re
import base64
from typing import Dict, Any, Optional

import nbformat
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ── Shading & borders ─────────────────────────────────────────────────────────

def _set_para_shading(para, fill_hex: str):
    pPr = para._p.get_or_add_pPr()
    for ex in pPr.findall(qn("w:shd")):
        pPr.remove(ex)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex.lstrip("#"))
    pPr.append(shd)


def _add_table_borders(table):
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement("w:tcBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                border = OxmlElement(f"w:{side}")
                border.set(qn("w:val"), "single")
                border.set(qn("w:sz"), "4")
                border.set(qn("w:space"), "0")
                border.set(qn("w:color"), "AAAAAA")
                tcBorders.append(border)
            tcPr.append(tcBorders)


def _shade_row(row, fill_hex: str):
    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill_hex.lstrip("#"))
        tcPr.append(shd)


# ── TOC field ─────────────────────────────────────────────────────────────────

def _insert_toc(doc: Document):
    """Insert a Table of Contents field at the current position (requires Word to update on open)."""
    para = doc.add_paragraph()
    run = para.add_run()

    fldBegin = OxmlElement("w:fldChar")
    fldBegin.set(qn("w:fldCharType"), "begin")
    run._r.append(fldBegin)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    run._r.append(instr)

    fldEnd = OxmlElement("w:fldChar")
    fldEnd.set(qn("w:fldCharType"), "end")
    run._r.append(fldEnd)

    note = doc.add_paragraph()
    note_run = note.add_run("(Table of Contents — press Ctrl+A then F9 in Word to update)")
    note_run.font.size = Pt(9)
    note_run.font.italic = True
    note_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    doc.add_paragraph()


# ── Inline markdown ───────────────────────────────────────────────────────────

def _apply_inline_markdown(text: str, para):
    parts = re.split(r"(\*\*.*?\*\*|\*.*?\*|`.*?`)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            r = para.add_run(part[2:-2])
            r.bold = True
        elif part.startswith("*") and part.endswith("*"):
            r = para.add_run(part[1:-1])
            r.italic = True
        elif part.startswith("`") and part.endswith("`"):
            r = para.add_run(part[1:-1])
            r.font.name = "Courier New"
            r.font.size = Pt(10)
        else:
            para.add_run(part)


# ── Markdown cell renderer ────────────────────────────────────────────────────

def _process_markdown(doc: Document, source: str):
    lines = source.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]

        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            doc.add_heading(m.group(2), level=min(len(m.group(1)), 4))
            i += 1
            continue

        if re.match(r"^[-*+]\s+", line):
            para = doc.add_paragraph(style="List Bullet")
            _apply_inline_markdown(re.sub(r"^[-*+]\s+", "", line), para)
            i += 1
            continue

        if re.match(r"^\d+\.\s+", line):
            para = doc.add_paragraph(style="List Number")
            _apply_inline_markdown(re.sub(r"^\d+\.\s+", "", line), para)
            i += 1
            continue

        if not line.strip():
            i += 1
            continue

        para = doc.add_paragraph()
        _apply_inline_markdown(line, para)
        i += 1


# ── Code cell ─────────────────────────────────────────────────────────────────

def _add_code_cell(doc: Document, source: str, cell_number: int):
    label = doc.add_paragraph()
    lr = label.add_run(f"[ Code Cell {cell_number} ]")
    lr.bold = True
    lr.italic = True
    lr.font.size = Pt(9)
    lr.font.color.rgb = RGBColor(0x77, 0x77, 0x77)
    label.paragraph_format.space_after = Pt(0)

    for line in (source or "").split("\n"):
        para = doc.add_paragraph()
        run = para.add_run(line)
        run.font.name = "Courier New"
        run.font.size = Pt(9)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        _set_para_shading(para, "F2F2F2")

    doc.add_paragraph()


# ── Text output (light yellow background) ─────────────────────────────────────

def _add_text_output(doc: Document, text: str):
    for line in text.split("\n"):
        if not line.strip():
            continue
        para = doc.add_paragraph()
        run = para.add_run(line)
        run.font.name = "Courier New"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        _set_para_shading(para, "FFFBE6")
    doc.add_paragraph()


# ── Image output ──────────────────────────────────────────────────────────────

def _add_image_output(doc: Document, b64_data: str, caption: Optional[str] = None):
    try:
        img_bytes = base64.b64decode(b64_data)
        doc.add_picture(io.BytesIO(img_bytes), width=Cm(14))
        if caption:
            cap = doc.add_paragraph()
            cap_run = cap.add_run(caption)
            cap_run.font.size = Pt(9)
            cap_run.italic = True
            cap_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
    except Exception:
        doc.add_paragraph("[Image output could not be embedded]")


# ── DataFrame table ───────────────────────────────────────────────────────────

def _add_dataframe_table(doc: Document, text: str):
    lines = [l for l in text.strip().split("\n") if l.strip()]
    if len(lines) < 2:
        _add_text_output(doc, text)
        return

    rows = [re.split(r"\s{2,}|\t", l.strip()) for l in lines]
    max_cols = max(len(r) for r in rows)
    if max_cols < 2:
        _add_text_output(doc, text)
        return

    table = doc.add_table(rows=len(rows), cols=max_cols)
    _add_table_borders(table)

    for ri, row_data in enumerate(rows):
        if ri % 2 == 1:
            _shade_row(table.rows[ri], "F7F7F7")
        for ci, cell_text in enumerate(row_data):
            cell = table.rows[ri].cells[ci]
            cell.text = cell_text.strip()
            if ri == 0:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True

    doc.add_paragraph()


# ── Main converter ────────────────────────────────────────────────────────────

def convert_notebook(file_bytes: bytes, cfg: Dict[str, Any]) -> bytes:
    nb = nbformat.reads(file_bytes.decode("utf-8"), as_version=4)
    doc = Document()

    raw_margins = cfg.get("margins", 2.5)
    if isinstance(raw_margins, dict):
        margins_val = float(raw_margins.get("top", 2.5))
    else:
        margins_val = float(raw_margins)

    for section in doc.sections:
        section.top_margin = Cm(margins_val)
        section.bottom_margin = Cm(margins_val)
        section.left_margin = Cm(margins_val)
        section.right_margin = Cm(margins_val)

    # Notebook title
    nb_title = (
        nb.get("metadata", {}).get("kernelspec", {}).get("display_name")
        or cfg.get("title")
        or "Notebook"
    )
    doc.add_heading(nb_title, level=1)

    # TOC
    add_toc = bool(cfg.get("addToc", False))
    if add_toc:
        doc.add_heading("Table of Contents", level=2)
        _insert_toc(doc)

    add_captions = bool(cfg.get("addCaptions", False))
    code_cell_count = 0
    figure_count = 0

    for cell in nb.cells:
        source = cell.get("source", "") or ""
        cell_type = cell.get("cell_type", "")

        if cell_type == "markdown":
            if source.strip():
                _process_markdown(doc, source)

        elif cell_type == "code":
            code_cell_count += 1
            if source.strip():
                _add_code_cell(doc, source, code_cell_count)

            for output in cell.get("outputs", []):
                otype = output.get("output_type", "")

                if otype == "stream":
                    text = "".join(output.get("text", []))
                    if text.strip():
                        _add_text_output(doc, text)

                elif otype in ("display_data", "execute_result"):
                    data = output.get("data", {})

                    embedded_image = False
                    for mime in ("image/png", "image/jpeg", "image/gif"):
                        if mime in data:
                            figure_count += 1
                            caption = f"Figure {figure_count}" if add_captions else None
                            _add_image_output(doc, data[mime], caption)
                            embedded_image = True
                            break

                    if not embedded_image and "text/plain" in data:
                        text = "".join(data["text/plain"])
                        if text.strip():
                            if re.search(r"\s{2,}", text) and "\n" in text:
                                _add_dataframe_table(doc, text)
                            else:
                                _add_text_output(doc, text)

                elif otype == "error":
                    ename = output.get("ename", "Error")
                    evalue = output.get("evalue", "")
                    para = doc.add_paragraph()
                    run = para.add_run(f"{ename}: {evalue}")
                    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
                    run.font.name = "Courier New"
                    run.font.size = Pt(9)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
