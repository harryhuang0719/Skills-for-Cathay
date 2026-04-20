"""
Cathay PPT Template — Constants
================================
All layout, brand, and typography constants in one place.
Imported by text_engine.py via `from constants import *`.
"""

import os

from pptx.util import Mm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


# ============================================================================
# 1. TEMPLATE PATH
# ============================================================================

TEMPLATE = os.path.expanduser("~/.claude/skills/cathay-ppt-template/assets/template.pptx")


# ============================================================================
# 2. BRAND COLORS
# ============================================================================

# ── Primary palette (一红主导 + 金点缀 + 灰分级) ──
CATHAY_RED       = RGBColor(0x80, 0x00, 0x00)   # 主色 MAROON (60-70% visual weight)
CATHAY_DARK_RED  = RGBColor(0x5E, 0x00, 0x00)   # 高对比 header (DARK_MAROON)
CATHAY_GOLD      = RGBColor(0xE8, 0xB0, 0x12)   # 强调色 金 (badge, accent bar)
CATHAY_LTGOLD    = RGBColor(0xF6, 0xDC, 0x92)   # 封面副文本 / 浅金 (PALE_GOLD)
CATHAY_ACCENT    = RGBColor(0xE6, 0x00, 0x00)   # 强调红 (warnings, emphasis)
CATHAY_PINK      = RGBColor(0xFF, 0x89, 0x89)   # 次级红 (bubble, risk chips)
CATHAY_SOFT_PINK = RGBColor(0xFE, 0xD3, 0xD3)   # 最浅红底 (panel bg)

# ── Text & Neutral ──
CATHAY_BLACK     = RGBColor(0x1A, 0x1A, 0x1A)   # 正文黑 (softer than pure black)
CATHAY_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
CATHAY_DARK_GREY = RGBColor(0x59, 0x59, 0x59)   # 次级文本
CATHAY_GREY      = RGBColor(0x80, 0x80, 0x80)   # 辅助 / 引用 (MID_GRAY)
CATHAY_LTGREY    = RGBColor(0xD9, 0xD9, 0xD9)   # table border / alt rows
CATHAY_LIGHT_BG  = RGBColor(0xF2, 0xF2, 0xF2)   # 面板底色 (LIGHT_GRAY)
CATHAY_VERY_LIGHT= RGBColor(0xFA, 0xFA, 0xFA)   # 更浅面板底

# Matplotlib chart palette
CATHAY_COLORS = ['#800000', '#E8B012', '#808080', '#E60000', '#F6DC92', '#404040', '#D9D9D9']

# ── Accent Pairs (color, light_background) for multi-item layouts ──
CATHAY_ACCENT_PAIRS = [
    (CATHAY_RED,    CATHAY_SOFT_PINK),  # maroon + soft pink bg
    (CATHAY_GOLD,   CATHAY_LTGOLD),     # gold + pale gold bg
    (CATHAY_ACCENT, CATHAY_PINK),       # accent red + pink bg
    (CATHAY_GREY,   CATHAY_LIGHT_BG),   # grey + light grey bg
]


# ============================================================================
# 3. TEXTBOX & PARAGRAPH DEFAULTS
# ============================================================================

MARGIN_ALL = Mm(2)  # 0.2cm all sides

DEFAULT_FONT_SIZE = 10       # pt — body text
INDENT_LEFT       = Mm(5)    # 0.5cm
SPACING_BEFORE    = Pt(4)
SPACING_AFTER     = Pt(0)
LINE_SPACING_PCT  = 120000   # 1.2x (IC memo standard)


# ============================================================================
# 4. CONTENT ZONE (mm)
# ============================================================================

CT = 31       # content top
CB = 181      # content bottom
CL = 11       # content left
CW = 233      # content width

CH  = CB - CT                          # 150mm content height
GAP_H = 5                             # horizontal gap (mm)
GAP_V = 3                             # vertical gap (mm)

# Grid layout constants
FULL   = CW
HALF   = (CW - GAP_H) / 2
THIRD  = (CW - GAP_H * 2) / 3
QUARTER = (CW - GAP_H * 3) / 4
ONE_THIRD    = (CW - GAP_H) * 1 / 3
TWO_THIRDS   = (CW - GAP_H) * 2 / 3
ONE_QUARTER  = (CW - GAP_H) * 1 / 4
THREE_QUARTER = (CW - GAP_H) * 3 / 4

# Column X positions (mm)
X1 = CL
X2_HALF  = CL + HALF + GAP_H
X2_Q34   = CL + ONE_QUARTER + GAP_H
X2_T23   = CL + ONE_THIRD + GAP_H
X2_MID   = CL + THIRD + GAP_H
X3_RIGHT = CL + THIRD * 2 + GAP_H * 2

# Row heights (mm)
ROW_FULL  = CH
ROW_HALF  = (CH - GAP_V) / 2
ROW_THIRD = (CH - GAP_V * 2) / 3

# Row Y positions (mm)
Y1 = CT
Y2_HALF = CT + ROW_HALF + GAP_V
Y2_MID  = CT + ROW_THIRD + GAP_V
Y3_BOT  = CT + ROW_THIRD * 2 + GAP_V * 2


# ============================================================================
# 5. SOURCE FOOTER CONSTANTS
# ============================================================================

SOURCE_FONT_PT       = 7
SOURCE_BOX_HEIGHT_MM = 5
SOURCE_Y_MM          = 182


# ============================================================================
# 6. CONTENT BOTTOM (for safe_textbox)
# ============================================================================

CONTENT_BOTTOM_MM = 175


# ============================================================================
# 7. LEGACY INCH-BASED CONSTANTS (backward compat)
# ============================================================================

CONTENT_LEFT  = 1.0 / 2.54
CONTENT_TOP   = 2.92 / 2.54
CONTENT_WIDTH = 23.4 / 2.54
CONTENT_HEIGHT = 14.6 / 2.54
SOURCE_LEFT   = 1.0 / 2.54
SOURCE_TOP    = 18.0 / 2.54
CONTENT_LEFT_CM = 1.0
CONTENT_TOP_CM  = 2.92


# ============================================================================
# 8. SECTION ICON CONSTANTS
# ============================================================================

ICON_FINANCIAL = (MSO_SHAPE.ROUNDED_RECTANGLE, 'E8B012')
ICON_INSIGHT   = (MSO_SHAPE.OVAL, '800000')
ICON_RISK      = (MSO_SHAPE.ISOSCELES_TRIANGLE, 'E60000')
ICON_CATALYST  = (MSO_SHAPE.DIAMOND, 'E8B012')
ICON_ACTION    = (MSO_SHAPE.RIGHT_ARROW, '800000')

_ICON_KEYWORD_MAP = {
    ICON_FINANCIAL: ["收入", "利润", "EPS", "现金流", "季度", "财务", "Revenue", "Margin", "Cash", "资产"],
    ICON_RISK:      ["风险", "Bear", "威胁", "下行", "Risk", "Kill", "止损"],
    ICON_CATALYST:  ["催化", "时间表", "Catalyst", "Trigger", "监控", "Monitoring"],
    ICON_ACTION:    ["行动", "决策", "建议", "Action", "Plan", "CIO", "裁决"],
    ICON_INSIGHT:   ["论点", "Thesis", "Bull", "洞察", "优势", "Moat", "Industry"],
}


# ============================================================================
# 9. CJK WIDTH TABLES
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
# 10. CJK CHARACTER DENSITY LIMITS (McKinsey pattern)
# ============================================================================

# Borrowed from McKinsey pattern — prevents text overflow in constrained boxes
CHAR_DENSITY_LIMITS = {
    5: 15, 10: 40, 15: 80, 20: 130, 25: 190,
    30: 260, 40: 450, 50: 700, 60: 1000, 75: 1200,
}


# ============================================================================
# 11. FONT SIZE GUARD RAILS
# ============================================================================

MIN_TITLE_FONT_PT = 18
MIN_BODY_FONT_PT = 9
MIN_SMALL_FONT_PT = 8
MIN_SOURCE_FONT_PT = 7


# ============================================================================
# 12. LAYOUT VARIETY RULE
# ============================================================================

MIN_UNIQUE_GRIDS_PER_25_SLIDES = 5
MAX_CONSECUTIVE_SAME_LAYOUT = 2
