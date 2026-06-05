"""
python-docx formatting logic for WordForge.
Applies typography, header/footer, margin, and spacing settings to a .docx.
"""

import io
from typing import Optional, Dict, Any

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re


# ── Helpers ───────────────────────────────────────────────────────────────────

def _hex_to_rgb(hex_color: str) -> Optional[RGBColor]:
    hex_color = hex_color.lstrip("#")
    if len(hex_color) == 6:
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        return RGBColor(r, g, b)
    return None


def _set_cell_border(cell, border_style="single", size=4):
    """Add borders to a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), border_style)
        border.set(qn('w:sz'), str(size))
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tcBorders.append(border)
    tcPr.append(tcBorders)


LINE_SPACING_MAP = {
    "1.0": 1.0,
    "1.15": 1.15,
    "1.5": 1.5,
    "2.0": 2.0,
}

STYLE_HEADING_MAP = {
    "h1": "Heading 1",
    "h2": "Heading 2",
    "h3": "Heading 3",
    "h4": "Heading 4",
}


# ── Core formatting ───────────────────────────────────────────────────────────

def apply_formatting(file_bytes: Optional[bytes], cfg: Dict[str, Any]) -> bytes:
    """
    Apply formatting settings to a .docx.
    If file_bytes is None, creates a sample document with placeholder content.
    """
    if file_bytes:
        doc = Document(io.BytesIO(file_bytes))
    else:
        doc = _create_sample_document()

    font_family = cfg.get("fontFamily", "Times New Roman")
    body_size = float(cfg.get("bodySize", 12))
    line_spacing = cfg.get("lineSpacing", "1.5")
    para_before = float(cfg.get("paraBefore", 0))
    para_after = float(cfg.get("paraAfter", 6))
    body_color = cfg.get("bodyColor", "#000000")

    heading_sizes = cfg.get("headingSizes", {
        "h1": 16, "h2": 14, "h3": 12, "h4": 11
    })
    heading_colors = cfg.get("headingColors", {})

    margins = cfg.get("margins", {
        "top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5
    })

    # Apply margins
    for section in doc.sections:
        section.top_margin = Cm(float(margins.get("top", 2.5)))
        section.bottom_margin = Cm(float(margins.get("bottom", 2.5)))
        section.left_margin = Cm(float(margins.get("left", 2.5)))
        section.right_margin = Cm(float(margins.get("right", 2.5)))

    # Apply header / footer
    header_cfg = cfg.get("header", {})
    footer_cfg = cfg.get("footer", {})

    for section in doc.sections:
        section.different_first_page_header_footer = bool(
            header_cfg.get("differentFirstPage") or footer_cfg.get("differentFirstPage")
        )

        # Header
        if header_cfg.get("text"):
            _set_header_footer_text(section.header, header_cfg["text"])
        if footer_cfg.get("text"):
            _set_header_footer_text(section.footer, footer_cfg["text"])

        # Page numbers
        if footer_cfg.get("pageNumbers"):
            _add_page_number(section.footer, footer_cfg.get("pageNumberPosition", "center"))

    # Apply paragraph styles
    spacing_val = LINE_SPACING_MAP.get(str(line_spacing), 1.5)

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ""

        is_heading = False
        for key, style in STYLE_HEADING_MAP.items():
            if style_name == style:
                is_heading = True
                size = float(heading_sizes.get(key, body_size))
                color_hex = heading_colors.get(key)

                for run in para.runs:
                    run.font.name = font_family
                    run.font.size = Pt(size)
                    if color_hex:
                        rgb = _hex_to_rgb(color_hex)
                        if rgb:
                            run.font.color.rgb = rgb
                break

        if not is_heading:
            for run in para.runs:
                run.font.name = font_family
                run.font.size = Pt(body_size)
                rgb = _hex_to_rgb(body_color)
                if rgb:
                    run.font.color.rgb = rgb

        # Line and paragraph spacing
        pf = para.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = spacing_val
        pf.space_before = Pt(para_before)
        pf.space_after = Pt(para_after)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _set_header_footer_text(hf, text: str):
    """Set plain text in a header or footer."""
    hf.paragraphs[0].clear()
    run = hf.paragraphs[0].add_run(text)
    run.font.size = Pt(10)


def _add_page_number(footer, position: str = "center"):
    """Add automatic page number field to footer."""
    para = footer.paragraphs[0]
    para.clear()

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    para.alignment = align_map.get(position, WD_ALIGN_PARAGRAPH.CENTER)

    run = para.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar)

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'PAGE'
    run._r.append(instrText)

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)


def _create_sample_document() -> Document:
    """Create a sample document with placeholder content for fresh-start formatting."""
    doc = Document()

    doc.add_heading("Document Title", level=1)
    doc.add_heading("Section Heading", level=2)
    doc.add_paragraph(
        "This is a sample paragraph. Replace this content with your own text. "
        "WordForge will apply your chosen formatting settings to the entire document."
    )
    doc.add_heading("Subsection", level=3)
    doc.add_paragraph(
        "Another paragraph of body text. The formatting controls on the left panel "
        "will affect font, size, spacing, margins, and heading styles throughout."
    )

    return doc
