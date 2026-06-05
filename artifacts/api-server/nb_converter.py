"""
.ipynb → .docx conversion logic for WordForge.
Handles markdown cells, code cells, output cells, images, and DataFrames.
"""

import io
import re
import base64
import json
from typing import Dict, Any, List

import nbformat
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    import mistletoe
    from mistletoe import Document as MdDocument
    from mistletoe.ast_renderer import AstRenderer
    HAS_MISTLETOE = True
except ImportError:
    HAS_MISTLETOE = False

try:
    from PIL import Image as PILImage
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


# ── Helpers ───────────────────────────────────────────────────────────────────

def _set_cell_shading(para, fill_color: str = "F2F2F2"):
    """Add light gray background shading to a paragraph."""
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    pPr.append(shd)


def _add_table_borders(table):
    """Add borders to all cells in a table."""
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                border = OxmlElement(f'w:{side}')
                border.set(qn('w:val'), 'single')
                border.set(qn('w:sz'), '4')
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), 'AAAAAA')
                tcBorders.append(border)
            tcPr.append(tcBorders)


def _apply_inline_markdown(run_text: str, para):
    """Apply bold/italic/code inline markdown in a paragraph."""
    # Simple inline: **bold**, *italic*, `code`
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*|`.*?`)', run_text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            r = para.add_run(part[2:-2])
            r.bold = True
        elif part.startswith('*') and part.endswith('*'):
            r = para.add_run(part[1:-1])
            r.italic = True
        elif part.startswith('`') and part.endswith('`'):
            r = para.add_run(part[1:-1])
            r.font.name = 'Courier New'
            r.font.size = Pt(10)
        else:
            para.add_run(part)


def _process_markdown_source(doc: Document, source: str):
    """Convert a markdown cell source into Word paragraphs."""
    lines = source.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]

        # ATX headings
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            level = len(m.group(1))
            text = m.group(2)
            doc.add_heading(text, level=min(level, 4))
            i += 1
            continue

        # Unordered list
        m = re.match(r'^[-*+]\s+(.*)', line)
        if m:
            para = doc.add_paragraph(style='List Bullet')
            _apply_inline_markdown(m.group(1), para)
            i += 1
            continue

        # Ordered list
        m = re.match(r'^\d+\.\s+(.*)', line)
        if m:
            para = doc.add_paragraph(style='List Number')
            _apply_inline_markdown(m.group(1), para)
            i += 1
            continue

        # Blank line → spacing
        if line.strip() == '':
            i += 1
            continue

        # Regular paragraph
        para = doc.add_paragraph()
        _apply_inline_markdown(line, para)
        i += 1


def _add_code_cell(doc: Document, source: str, cell_number: int):
    """Add a code cell as monospaced shaded paragraphs."""
    label = doc.add_paragraph()
    label_run = label.add_run(f"[Code Cell {cell_number}]")
    label_run.bold = True
    label_run.font.size = Pt(9)
    label_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    for line in (source or '').split('\n'):
        para = doc.add_paragraph()
        run = para.add_run(line)
        run.font.name = 'Courier New'
        run.font.size = Pt(9)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        _set_cell_shading(para, 'F2F2F2')


def _add_text_output(doc: Document, text: str):
    """Add plain text output."""
    for line in text.split('\n'):
        if not line.strip():
            continue
        para = doc.add_paragraph()
        run = para.add_run(line)
        run.font.name = 'Courier New'
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        _set_cell_shading(para, 'FAFAFA')


def _add_image_output(doc: Document, b64_data: str, mime: str):
    """Embed a base64-encoded image into the document."""
    try:
        img_bytes = base64.b64decode(b64_data)
        stream = io.BytesIO(img_bytes)
        # Try to add image; python-docx auto-sizes
        doc.add_picture(stream, width=Cm(14))
        doc.add_paragraph()  # spacing after image
    except Exception:
        doc.add_paragraph("[Image output could not be embedded]")


def _add_dataframe_table(doc: Document, html_or_text: str):
    """Attempt to render a DataFrame output as a Word table."""
    # Try to parse pipe-separated or simple text table
    lines = [l for l in html_or_text.strip().split('\n') if l.strip()]
    if len(lines) < 2:
        _add_text_output(doc, html_or_text)
        return

    # Check if it looks like a text table (has spaces/tabs as separators)
    rows = []
    for line in lines:
        # Split on 2+ spaces or tabs
        cells = re.split(r'\s{2,}|\t', line.strip())
        if cells:
            rows.append(cells)

    if not rows:
        _add_text_output(doc, html_or_text)
        return

    max_cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=max_cols)
    _add_table_borders(table)

    for ri, row_data in enumerate(rows):
        for ci, cell_text in enumerate(row_data):
            cell = table.rows[ri].cells[ci]
            cell.text = cell_text.strip()
            if ri == 0:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True

    doc.add_paragraph()  # spacing after table


# ── Main converter ────────────────────────────────────────────────────────────

def convert_notebook(file_bytes: bytes, cfg: Dict[str, Any]) -> bytes:
    """Convert a Jupyter notebook (.ipynb) to a formatted .docx."""
    nb = nbformat.reads(file_bytes.decode('utf-8'), as_version=4)

    doc = Document()

    # Apply basic formatting from cfg
    margins = cfg.get("margins", {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5})
    for section in doc.sections:
        section.top_margin = Cm(float(margins.get("top", 2.5)))
        section.bottom_margin = Cm(float(margins.get("bottom", 2.5)))
        section.left_margin = Cm(float(margins.get("left", 2.5)))
        section.right_margin = Cm(float(margins.get("right", 2.5)))

    # Title from notebook metadata
    nb_name = nb.get('metadata', {}).get('kernelspec', {}).get('display_name', 'Notebook')
    doc.add_heading(nb_name, level=1)

    code_cell_count = 0

    for cell in nb.cells:
        source = cell.get('source', '') or ''
        cell_type = cell.get('cell_type', '')

        if cell_type == 'markdown':
            if source.strip():
                _process_markdown_source(doc, source)

        elif cell_type == 'code':
            code_cell_count += 1
            if source.strip():
                _add_code_cell(doc, source, code_cell_count)

            # Process outputs
            for output in cell.get('outputs', []):
                output_type = output.get('output_type', '')

                if output_type == 'stream':
                    text = ''.join(output.get('text', []))
                    if text.strip():
                        _add_text_output(doc, text)

                elif output_type in ('display_data', 'execute_result'):
                    data = output.get('data', {})

                    # Images first
                    for mime in ('image/png', 'image/jpeg', 'image/gif'):
                        if mime in data:
                            _add_image_output(doc, data[mime], mime)
                            break
                    else:
                        # Check for DataFrame-like text/plain
                        if 'text/plain' in data:
                            text = ''.join(data['text/plain'])
                            if text.strip():
                                # Heuristic: if it looks tabular, try table
                                if re.search(r'\s{2,}', text) and '\n' in text:
                                    _add_dataframe_table(doc, text)
                                else:
                                    _add_text_output(doc, text)

                elif output_type == 'error':
                    ename = output.get('ename', 'Error')
                    evalue = output.get('evalue', '')
                    para = doc.add_paragraph()
                    run = para.add_run(f"{ename}: {evalue}")
                    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
                    run.font.name = 'Courier New'
                    run.font.size = Pt(9)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
