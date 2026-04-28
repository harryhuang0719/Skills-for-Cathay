"""
Cathay PPT Template v3 — Constants
====================================
Single source of truth. Based on 阿维塔 template (10.00" x 7.50", 4:3).

Key layout: left dark red vertical line (5mm wide, full height) on Layout [4].
Content zone starts to the right of the red line.
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

CATHAY_RED       = RGBColor(0x80, 0x00, 0x00)   # 主色 MAROON
CATHAY_DARK_RED  = RGBColor(0x5E, 0x00, 0x00)   # 高对比 header
CATHAY_GOLD      = RGBColor(0xE8, 0xB0, 0x12)   # 强调金
CATHAY_LTGOLD    = RGBColor(0xF6, 0xDC, 0x92)   # 浅金
CATHAY_ACCENT    = RGBColor(0xE6, 0x00, 0x00)   # 强调红
CATHAY_PINK      = RGBColor(0xFF, 0x89, 0x89)   # 次级红
CATHAY_SOFT_PINK = RGBColor(0xFE, 0xD3, 0xD3)   # 最浅红底

CATHAY_BLACK     = RGBColor(0x1A, 0x1A, 0x1A)   # 正文黑
CATHAY_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
CATHAY_DARK_GREY = RGBColor(0x59, 0x59, 0x59)   # 次级文本
CATHAY_GREY      = RGBColor(0x80, 0x80, 0x80)   # 辅助
CATHAY_LTGREY    = RGBColor(0xD9, 0xD9, 0xD9)   # table borders
CATHAY_LIGHT_BG  = RGBColor(0xF2, 0xF2, 0xF2)   # 面板底色
CATHAY_VERY_LIGHT= RGBColor(0xFA, 0xFA, 0xFA)   # 更浅面板底

# Matplotlib chart palette
CATHAY_COLORS = ['#800000', '#E8B012', '#808080', '#E60000', '#F6DC92', '#404040', '#D9D9D9']

CATHAY_ACCENT_PAIRS = [
    (CATHAY_RED,    CATHAY_SOFT_PINK),
    (CATHAY_GOLD,   CATHAY_LTGOLD),
    (CATHAY_ACCENT, CATHAY_PINK),
    (CATHAY_GREY,   CATHAY_LIGHT_BG),
]


# ============================================================================
# 3. TYPOGRAPHY (v3 — calibrated to 阿维塔 reference)
# ============================================================================

MARGIN_ALL      = Mm(2)    # 0.2cm all sides
DEFAULT_FONT_SIZE = 11.5   # pt — body text (up from 10pt)
INDENT_LEFT     = Mm(5)    # 0.5cm
SPACING_BEFORE  = Pt(5)    # 5pt before (up from 4pt)
SPACING_AFTER   = Pt(0)
LINE_SPACING_PCT = 130000  # 1.3x (up from 1.2x) — more breathing room

# Font size hierarchy
TITLE_FONT_PT     = 20   # page title, dark red bold
SUBTITLE_FONT_PT  = 14   # section subtitle
BODY_FONT_PT      = 11.5 # bullet body
SMALL_FONT_PT     = 10   # table cells, sub-bullets
CAPTION_FONT_PT   = 8    # source footer, chart labels
KPI_VALUE_PT      = 18   # KPI number
KPI_LABEL_PT      = 8    # KPI label

# Chinese font: STKaiti (华文楷体) — matches 阿维塔 reference
# English font: Calibri
CN_FONT = "STKaiti"
EN_FONT = "Calibri"


# ============================================================================
# 4. SPACING TIERS
# ============================================================================

GAP_XS = 2   # header internal padding
GAP_SM = 4   # same-group element spacing
GAP_MD = 6   # cross-section spacing
GAP_LG = 10  # major breathing room

GAP_H = GAP_MD  # column gap
GAP_V = GAP_SM  # row gap


# ============================================================================
# 5. CONTENT ZONE (mm) — accounts for left red line (5mm)
# ============================================================================

# Layout [4] has a left red vertical line "Rectangle 9" at x=0, w=5mm, full height
RED_LINE_WIDTH = 5     # mm, built into Layout [4]
CL = 10                # content left — 5mm line + 5mm spacing (matches template title at 10.6mm)
CT = 31                # content top (below title zone)
CB = 181               # content bottom (above source footer)
CW = 234               # content width (254 - 10 - 10)

CH  = CB - CT          # 150mm content height
CONTENT_BOTTOM_MM = 175

# Grid layout constants
FULL   = CW                         # 234
HALF   = (CW - GAP_H) / 2           # 114.0
THIRD  = (CW - GAP_H * 2) / 3       # 74.0
QUARTER = (CW - GAP_H * 3) / 4      # 54.0
ONE_THIRD    = (CW - GAP_H) * 1 / 3 # 76.0
TWO_THIRDS   = (CW - GAP_H) * 2 / 3 # 152.0
ONE_QUARTER  = (CW - GAP_H) * 1 / 4 # 57.0
THREE_QUARTER = (CW - GAP_H) * 3 / 4 # 171.0

# Column X positions (mm)
X1 = CL                             # 10
X2_HALF  = CL + HALF + GAP_H        # 130.0
X2_Q34   = CL + ONE_QUARTER + GAP_H # 73.0
X2_T23   = CL + ONE_THIRD + GAP_H   # 92.0
X2_MID   = CL + THIRD + GAP_H       # 90.0
X3_RIGHT = CL + THIRD * 2 + GAP_H * 2  # 170.0

# Row heights (mm)
ROW_FULL  = CH                       # 150
ROW_HALF  = (CH - GAP_V) / 2        # 73.0
ROW_THIRD = (CH - GAP_V * 2) / 3    # ~47.33

# Row Y positions (mm)
Y1 = CT                             # 31
Y2_HALF = CT + ROW_HALF + GAP_V     # 108.0
Y2_MID  = CT + ROW_THIRD + GAP_V    # ~82.33
Y3_BOT  = CT + ROW_THIRD * 2 + GAP_V * 2  # ~133.67


# ============================================================================
# 6. SOURCE FOOTER
# ============================================================================

SOURCE_FONT_PT       = 8  # up from 7pt
SOURCE_BOX_HEIGHT_MM = 5
SOURCE_Y_MM          = 182


# ============================================================================
# 7. SECTION ICON CONSTANTS
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
# 8. CJK CHARACTER DENSITY LIMITS
# ============================================================================

CHAR_DENSITY_LIMITS = {
    5: 15, 10: 40, 15: 80, 20: 130, 25: 190,
    30: 260, 40: 450, 50: 700, 60: 1000, 75: 1200,
}


# ============================================================================
# 9. FONT SIZE GUARD RAILS
# ============================================================================

MIN_TITLE_FONT_PT = 18
MIN_BODY_FONT_PT = 9
MIN_SMALL_FONT_PT = 8
MIN_SOURCE_FONT_PT = 7


# ============================================================================
# 10. LAYOUT VARIETY
# ============================================================================

MIN_UNIQUE_GRIDS_PER_25_SLIDES = 5
MAX_CONSECUTIVE_SAME_LAYOUT = 2
