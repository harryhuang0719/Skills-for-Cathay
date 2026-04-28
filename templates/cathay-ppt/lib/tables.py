"""
Cathay PPT Template — Table Helpers
=====================================
add_table(), smart_table() with auto row-height calculation.

Usage:
    from tables import add_table, smart_table
"""

import math

from pptx.util import Pt, Mm
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

from constants import (
    CL, CT, CW,
    CATHAY_RED, CATHAY_WHITE, CATHAY_BLACK, CATHAY_LIGHT_BG,
)
from fonts import set_run_font, get_char_width


# ============================================================================
# 1. STANDARD TABLE
# ============================================================================

def add_table(slide, data, left_mm=None, top_mm=None, width_mm=None,
              row_height=None, font_size=None, col_widths=None):
    """Add a Cathay-formatted table with dark-red header and alternating rows.

    Args:
        slide: pptx slide object
        data: 2D list [[row0_col0, row0_col1, ...], [row1_col0, ...], ...]
              First row is treated as header.
        left_mm, top_mm, width_mm: position and size in mm
        row_height: height per row in mm (default 7)
        font_size: font size in pt (default 9)
        col_widths: list of mm widths per column (optional, equal if omitted)

    Returns:
        (table_object, bottom_y_mm)
    """
    left_mm = left_mm or CL
    top_mm = top_mm or CT
    width_mm = width_mm or CW
    row_height = row_height or 7
    font_size = font_size or 9

    rows = len(data)
    cols = len(data[0]) if data else 0
    if rows == 0 or cols == 0:
        return None, top_mm

    total_h = rows * row_height
    table_shape = slide.shapes.add_table(
        rows, cols,
        Mm(left_mm), Mm(top_mm),
        Mm(width_mm), Mm(total_h))
    table = table_shape.table

    # Set column widths
    if col_widths:
        for ci, cw in enumerate(col_widths):
            if ci < cols:
                table.columns[ci].width = Mm(cw)
    else:
        eq_w = width_mm / cols
        for ci in range(cols):
            table.columns[ci].width = Mm(eq_w)

    for i in range(rows):
        table.rows[i].height = Mm(row_height)
        for j in range(cols):
            cell = table.cell(i, j)
            cell.margin_left = Mm(1.5)
            cell.margin_right = Mm(1.5)
            cell.margin_top = Mm(1)
            cell.margin_bottom = Mm(1)

            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT

            cell_text = str(data[i][j]) if data[i][j] is not None else ""
            run = p.add_run()

            if i == 0:
                set_run_font(run, cell_text, size_pt=font_size, bold=True,
                             color_rgb=CATHAY_WHITE)
            else:
                set_run_font(run, cell_text, size_pt=font_size,
                             color_rgb=CATHAY_BLACK)

            # Background colors
            tcPr = cell._tc.get_or_add_tcPr()
            if i == 0:
                solidFill = tcPr.makeelement(qn('a:solidFill'), {})
                srgbClr = solidFill.makeelement(qn('a:srgbClr'), {'val': '800000'})
                solidFill.append(srgbClr)
                tcPr.append(solidFill)
            elif i % 2 == 0:
                solidFill = tcPr.makeelement(qn('a:solidFill'), {})
                srgbClr = solidFill.makeelement(qn('a:srgbClr'), {'val': 'F2F2F2'})
                solidFill.append(srgbClr)
                tcPr.append(solidFill)

    bottom_y = top_mm + total_h
    return table, bottom_y


# ============================================================================
# 2. SMART TABLE (auto-fit row heights)
# ============================================================================

def smart_table(slide, data, left_mm=None, top_mm=None, width_mm=None,
                max_bottom_mm=180, font_size=9, min_row_h=7):
    """Create auto-fit table: scan cell content, auto-set row heights.

    Returns:
        (table, bottom_y_mm)
    """
    left_mm = left_mm or CL
    top_mm = top_mm or CT
    width_mm = width_mm or CW
    rows = len(data)
    cols = len(data[0]) if data else 0

    col_w = width_mm / cols if cols > 0 else width_mm

    # Calculate row heights based on content
    row_heights = []
    for ri, row_data in enumerate(data):
        max_lines = 1
        for ci, cell_text in enumerate(row_data):
            text = str(cell_text) if cell_text else ""
            if not text:
                continue
            has_cjk = any('\u4e00' <= c <= '\u9fff' for c in text)
            char_w = get_char_width(font_size, has_cjk)
            usable = col_w - 3
            if usable <= 0:
                usable = 5
            lines = max(1, math.ceil(len(text) * char_w / usable))
            lines += text.count('\n')
            max_lines = max(max_lines, lines)

        line_h = font_size * 0.3528 * 1.2
        needed_h = max_lines * line_h + 3
        row_heights.append(max(needed_h, min_row_h))

    total_h = sum(row_heights)

    # Scale down font if table doesn't fit
    if top_mm + total_h > max_bottom_mm:
        for smaller_font in [font_size - 0.5, font_size - 1, font_size - 1.5, font_size - 2]:
            if smaller_font < 7:
                break
            font_size = smaller_font
            row_heights_new = []
            for ri, row_data in enumerate(data):
                max_lines = 1
                for ci, cell_text in enumerate(row_data):
                    text = str(cell_text) if cell_text else ""
                    if not text:
                        continue
                    has_cjk = any('\u4e00' <= c <= '\u9fff' for c in text)
                    char_w = get_char_width(smaller_font, has_cjk)
                    usable = col_w - 3
                    if usable <= 0:
                        usable = 5
                    lines = max(1, math.ceil(len(text) * char_w / usable))
                    max_lines = max(max_lines, lines)
                line_h = smaller_font * 0.3528 * 1.2
                needed_h = max_lines * line_h + 3
                row_heights_new.append(max(needed_h, 6))
            total_h_new = sum(row_heights_new)
            if top_mm + total_h_new <= max_bottom_mm:
                row_heights = row_heights_new
                total_h = total_h_new
                break

    avg_row_h = total_h / rows if rows > 0 else min_row_h
    return add_table(slide, data, left_mm=left_mm, top_mm=top_mm,
                     width_mm=width_mm, row_height=avg_row_h, font_size=font_size)


__all__ = [
    "add_table",
    "smart_table",
]
