"""
Cover page generator for WordForge.
Inserts a styled cover page as the first section of a document.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from typing import Dict, Any


def _set_run_color(run, hex_color: str):
    h = hex_color.lstrip("#").upper()
    try:
        run.font.color.rgb = RGBColor(
            int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        )
    except Exception:
        pass
    rPr = run._r.get_or_add_rPr()
    for ex in rPr.findall(qn("w:color")):
        rPr.remove(ex)
    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), h)
    rPr.append(color_el)


def add_cover_page(doc: Document, cover_cfg: Dict[str, Any]):
    """
    Prepend a cover page to the beginning of the document.
    Cover page fields: title, studentName, rollNumber, subject,
    department, universityName, date
    """
    title = cover_cfg.get("title", "Project Title")
    student_name = cover_cfg.get("studentName", "")
    roll_number = cover_cfg.get("rollNumber", "")
    subject = cover_cfg.get("subject", "")
    department = cover_cfg.get("department", "")
    university = cover_cfg.get("universityName", "")
    date = cover_cfg.get("date", "")

    # Insert paragraphs at the beginning (before existing content)
    # We insert a page break after the cover section
    cover_paras = []

    # Spacing at top
    p_space = OxmlElement("w:p")
    doc.element.body.insert(0, p_space)
    cover_paras.append(p_space)

    def _insert_para(text: str, font_size: float, bold: bool, center: bool,
                     color: str = "#000000", top_space: float = 0, bottom_space: float = 6,
                     index: int = 0):
        p = doc.paragraphs[0]._p  # reference point
        new_p = OxmlElement("w:p")

        pPr = OxmlElement("w:pPr")
        if center:
            jc = OxmlElement("w:jc")
            jc.set(qn("w:val"), "center")
            pPr.append(jc)
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"), str(int(top_space * 20)))
        spacing.set(qn("w:after"), str(int(bottom_space * 20)))
        pPr.append(spacing)
        new_p.append(pPr)

        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(font_size * 2)))
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(int(font_size * 2)))
        rPr.append(sz)
        rPr.append(szCs)

        if bold:
            b = OxmlElement("w:b")
            bCs = OxmlElement("w:bCs")
            rPr.append(b)
            rPr.append(bCs)

        h = color.lstrip("#").upper()
        c = OxmlElement("w:color")
        c.set(qn("w:val"), h)
        rPr.append(c)

        r.append(rPr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        new_p.append(r)

        return new_p

    # Build cover page paragraphs (in reverse order, inserting at position 0)
    # We'll collect them and then insert all in order

    cover_elements = []

    # University name at top
    if university:
        cover_elements.append(
            _insert_para(university, 14, True, True, "#000000", top_space=72, bottom_space=6)
        )

    # Department
    if department:
        cover_elements.append(
            _insert_para(department, 12, False, True, "#555555", bottom_space=48)
        )

    # Title — large and prominent
    cover_elements.append(
        _insert_para(title, 22, True, True, "#1F3864", top_space=36, bottom_space=12)
    )

    # Subject
    if subject:
        cover_elements.append(
            _insert_para(subject, 12, False, True, "#333333", top_space=24, bottom_space=6)
        )

    # Student name
    if student_name:
        cover_elements.append(
            _insert_para(f"Submitted by: {student_name}", 12, False, True, "#000000", top_space=48, bottom_space=4)
        )

    # Roll number
    if roll_number:
        cover_elements.append(
            _insert_para(f"Roll No: {roll_number}", 11, False, True, "#555555", bottom_space=4)
        )

    # Date
    if date:
        cover_elements.append(
            _insert_para(date, 11, False, True, "#555555", top_space=36, bottom_space=0)
        )

    # Page break after cover
    pb_p = OxmlElement("w:p")
    pb_r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    pb_r.append(br)
    pb_p.append(pb_r)
    cover_elements.append(pb_p)

    # Insert all cover elements at the very beginning of the document body
    body = doc.element.body
    for i, el in enumerate(reversed(cover_elements)):
        body.insert(0, el)

    return doc
