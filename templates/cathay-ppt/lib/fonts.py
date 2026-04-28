"""
Cathay PPT Template — Font Engine
===================================
set_run_font(), add_mixed_text(), get_char_width(), CJK/LATIN width tables.

Usage:
    from fonts import set_run_font, add_mixed_text, get_char_width
"""

import re

from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree

from constants import DEFAULT_FONT_SIZE


# ============================================================================
# 1. CJK/LATIN CHARACTER WIDTH TABLES (per font-size, in mm)
# ============================================================================

CJK_CHAR_WIDTH = {
    7:    2.2,
    7.5:  2.4,
    8:    2.5,
    8.5:  2.7,
    9:    2.85,
    9.5:  3.0,
    10:   3.15,
    10.5: 3.3,
    11:   3.5,
    12:   3.8,
    14:   4.4,
    16:   5.0,
    18:   5.7,
    20:   6.3,
    22:   6.9,
    24:   7.6,
    26:   8.2,
    28:   8.8,
}

LATIN_CHAR_WIDTH = {
    7:    1.4,
    7.5:  1.5,
    8:    1.6,
    8.5:  1.7,
    9:    1.8,
    9.5:  1.9,
    10:   2.0,
    10.5: 2.1,
    11:   2.2,
    12:   2.4,
    14:   2.8,
    16:   3.2,
    18:   3.6,
    20:   4.0,
    22:   4.4,
    24:   4.8,
    26:   5.2,
    28:   5.6,
}


# ============================================================================
# 2. CHARACTER WIDTH LOOKUP
# ============================================================================

def get_char_width(font_pt, is_cjk=False):
    """Get character width in mm for given font size."""
    table = CJK_CHAR_WIDTH if is_cjk else LATIN_CHAR_WIDTH
    pts = sorted(table.keys())
    if font_pt <= pts[0]:
        return table[pts[0]]
    if font_pt >= pts[-1]:
        return table[pts[-1]]
    for i in range(len(pts) - 1):
        if pts[i] <= font_pt <= pts[i + 1]:
            ratio = (font_pt - pts[i]) / (pts[i + 1] - pts[i])
            return table[pts[i]] + ratio * (table[pts[i + 1]] - table[pts[i]])
    return table[10]


# ============================================================================
# 3. FONT SETTING (with auto CJK/English detection)
# ============================================================================

def set_run_font(run, text, size_pt=None, bold=False, color_rgb=None):
    """Set font with auto Chinese/English detection. Default size: 10pt.

    Sets KaiTi for CJK text, Calibri for Latin text.
    Also sets a:ea typeface, a:altLang, and a:cs for cross-platform rendering.
    """
    size_pt = size_pt or DEFAULT_FONT_SIZE
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color_rgb:
        run.font.color.rgb = color_rgb

    # Extended CJK detection: CJK chars + CJK punctuation + fullwidth forms
    has_chinese = any(
        '\u4e00' <= c <= '\u9fff' or '\u3000' <= c <= '\u303f' or '\uff00' <= c <= '\uffef'
        for c in text
    )
    if has_chinese:
        run.font.name = "STKaiti"
        rPr = run._r.get_or_add_rPr()
        rPr.set(qn('a:altLang'), 'zh-CN')
        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = etree.SubElement(rPr, qn('a:ea'))
        ea.set('typeface', 'STKaiti')
        cs = rPr.find(qn('a:cs'))
        if cs is None:
            cs = etree.SubElement(rPr, qn('a:cs'))
        cs.set('typeface', 'STKaiti')
    else:
        run.font.name = "Calibri"
        rPr = run._r.get_or_add_rPr()
        cs = rPr.find(qn('a:cs'))
        if cs is None:
            cs = etree.SubElement(rPr, qn('a:cs'))
        cs.set('typeface', 'Calibri')


def add_mixed_text(para, text, size_pt=None, bold=False, color_rgb=None):
    """Split mixed CJK/Latin text into multiple runs, each with correct font.

    Use this for any paragraph that may contain both Chinese and English text.
    """
    size_pt = size_pt or DEFAULT_FONT_SIZE
    segments = re.findall(
        r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]+|[^\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]+',
        text
    )
    for seg in segments:
        if seg.strip() or seg == ' ':
            run = para.add_run()
            set_run_font(run, seg, size_pt=size_pt, bold=bold, color_rgb=color_rgb)


__all__ = [
    "CJK_CHAR_WIDTH",
    "LATIN_CHAR_WIDTH",
    "get_char_width",
    "set_run_font",
    "add_mixed_text",
]
