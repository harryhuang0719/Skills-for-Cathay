---
name: cathay-ppt-template
description: Cathay Capital / Smart Energy Fund branded PowerPoint template for creating investment memos, IC materials, pitch decks, and client presentations. Use when creating Cathay-branded PPT decks.
---

# Cathay PPT Template

Template: `assets/template.pptx` (10.00" x 7.50", 4:3, 12 layouts)

**MUST READ** (按顺序):
1. `references/ppt-generation-rules.md` — 全局PPT生成铁律（overflow prevention, bullet hierarchy, conclusion titles, font rules）
2. `references/text-fitting-engine.md` — 文字高度计算引擎、auto-fit textbox、smart table、merge函数（**2026-03-25新增**）

所有使用此模板的 skill 都必须遵守。

## Quick Start — Using the lib/ modules

All PPT generation should use the pre-built modules in `lib/`:

```python
import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/cathay-ppt-template/lib"))

# Constants (single source of truth — brand colors, grid, guard rails)
from constants import *

# Core engine (text fitting, shapes, validation, anti-corruption)
from text_engine import *

# Pre-built slide templates (16 layouts)
from slide_templates import *

# QC automation (guard rails, autofix pipeline, PDF export)
from qc_automation import full_qc_pipeline, check_guard_rails, autofix_pipeline, auto_fix_all

# Data-driven generation (specs -> deck)
from data_driven import build_deck_from_specs, DataRegistry
```

### Workflow
1. Define data in `DataRegistry` (single source of truth)
2. Define slides as specs (template + data)
3. Call `build_deck_from_specs(specs, output_path)`
4. Run `full_qc_pipeline(output_path)` to validate
5. Review PDF output

### Module Reference

| Module | File | Key Exports |
|--------|------|-------------|
| `constants` | `lib/constants.py` | **Single source of truth** for all brand colors, grid constants, accent pairs, font guard rails, CJK width tables, density limits |
| `text_engine` | `lib/text_engine.py` | `calc_text_height`, `calc_textframe_height`, `smart_textbox`, `smart_table`, `validate_and_fix`, `save_with_validation` (with anti-corruption), `full_cleanup`, `_clean_shape`, `merge_slides` |
| `slide_templates` | `lib/slide_templates.py` | **16 templates**: T1-T10 (original) + T11 `template_donut_chart` + T12 `template_before_after` + T13 `template_funnel` + T14 `template_swot` + T15 `template_waterfall` + T16 `template_stakeholder_map` |
| `qc_automation` | `lib/qc_automation.py` | `full_qc_pipeline` (with guard rails + autofix), `check_guard_rails` (8 rules), `autofix_pipeline` (4-stage), `update_slide_in_deck`, `batch_validate`, `auto_fix_all` |
| `data_driven` | `lib/data_driven.py` | `DataRegistry` (data+source tracking), `build_deck_from_specs` (specs->deck), `render_spec`, `TEMPLATE_ROUTER` |

## Brand Identity

| Element | Value |
|---------|-------|
| Primary Color | `#800000` MAROON — title bars, headers, callouts (60-70% visual weight) |
| Dark Red | `#5E0000` DARK_MAROON — high-contrast headers |
| Accent Gold | `#E8B012` — badge, accent bar, emphasis |
| Light Gold | `#F6DC92` PALE_GOLD — cover subtitle, light accent |
| Accent Red | `#E60000` — warnings, emphasis |
| Pink | `#FF8989` — risk chips, bubble fills |
| Soft Pink | `#FED3D3` — lightest red panel background |
| Body Text | `#1A1A1A` (soft black) — **not pure #000000** |
| White | `#FFFFFF` — **always use white on dark/red backgrounds** |
| Dark Grey | `#595959` — secondary text |
| Mid Grey | `#808080` — auxiliary / citations |
| Light Grey | `#D9D9D9` — table borders, alternating rows |
| Panel BG | `#F2F2F2` — panel / card background |
| Very Light BG | `#FAFAFA` — even lighter panels |
| English/Number Font | **Calibri** |
| Chinese Font | **楷体 (KaiTi)** |
| Title Size | 20-28pt, **left-aligned** |
| Subtitle | 14pt |
| Body Text | **10pt** (default for all content) |
| Source / Notes | 8pt, grey, bottom-left |
| Alignment | **Left-aligned** throughout (titles, cover, body) |
| TextBox Margins | **0.2cm** all four sides (left, right, top, bottom) |
| TextBox AutoFit | **Do NOT shrink text to fit** — use `MSO_AUTO_SIZE.NONE` |
| Paragraph Indent | Left **0.5cm** (for bulleted content) |
| Spacing Before | **4pt** |
| Spacing After | **0pt** |
| Line Spacing | **1.2 lines** (120%) — IC memos; 1.3 for pitch decks |
| Bullet Style | **Filled square** via PPT buChar XML (not text character) at 70% size |

## Global Style Constants

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Mm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_LINE_SPACING
from pptx.oxml.ns import qn
from lxml import etree

TEMPLATE = os.path.expanduser("~/.claude/skills/cathay-ppt-template/assets/template.pptx")

# ── Brand Colors (一红主导 + 金点缀 + 灰分级) ──
CATHAY_RED       = RGBColor(0x80, 0x00, 0x00)  # primary MAROON
CATHAY_DARK_RED  = RGBColor(0x5E, 0x00, 0x00)  # high-contrast header
CATHAY_GOLD      = RGBColor(0xE8, 0xB0, 0x12)  # accent gold
CATHAY_LTGOLD    = RGBColor(0xF6, 0xDC, 0x92)  # pale gold
CATHAY_ACCENT    = RGBColor(0xE6, 0x00, 0x00)  # accent red
CATHAY_PINK      = RGBColor(0xFF, 0x89, 0x89)  # risk chips
CATHAY_SOFT_PINK = RGBColor(0xFE, 0xD3, 0xD3)  # lightest red bg
CATHAY_BLACK     = RGBColor(0x1A, 0x1A, 0x1A)  # body text (soft black)
CATHAY_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)  # text on dark bg
CATHAY_DARK_GREY = RGBColor(0x59, 0x59, 0x59)  # secondary text
CATHAY_GREY      = RGBColor(0x80, 0x80, 0x80)  # auxiliary
CATHAY_LTGREY    = RGBColor(0xD9, 0xD9, 0xD9)  # table borders
CATHAY_LIGHT_BG  = RGBColor(0xF2, 0xF2, 0xF2)  # panel bg
CATHAY_VERY_LIGHT= RGBColor(0xFA, 0xFA, 0xFA)  # even lighter panels

# ── TextBox Internal Margins: 0.2cm all sides ──
MARGIN_ALL = Mm(2)  # 0.2cm = 2mm

# ── Paragraph Defaults ──
DEFAULT_FONT_SIZE = 10.5     # pt — default content font size
INDENT_LEFT       = Mm(5)   # 0.5cm left indent for bulleted paragraphs
SPACING_BEFORE    = Pt(4)   # 4pt before
SPACING_AFTER     = Pt(0)   # 0pt after
LINE_SPACING_PCT  = 120000  # 1.2 lines = 120% (IC memos)

# ── Content & Source Positioning (cm) ──
CONTENT_TOP_CM    = 2.92    # content text box / table default top
CONTENT_LEFT_CM   = 1.0     # content text box / table default left
SOURCE_LEFT_CM    = 1.0     # source footer left position
SOURCE_TOP_CM     = 18.0    # source footer top position

# Content placement bounds (for shapes/tables within slides)
# Content placement bounds (cm → mm for python-pptx Mm())
CONTENT_LEFT_MM   = 10      # 1.0cm
CONTENT_TOP_MM    = 29.2    # 2.92cm — below title bar
CONTENT_RIGHT_MM  = 244     # ~24.4cm (slide width 25.4 - ~1cm right margin)
CONTENT_WIDTH_MM  = 234     # CONTENT_RIGHT - CONTENT_LEFT
CONTENT_BOTTOM_MM = 175     # above source footer area

# Source footer position (cm → inches)
SOURCE_LEFT  = 1.0 / 2.54   # 1cm = 0.394"
SOURCE_TOP   = 18.0 / 2.54  # 18cm = 7.087"
```

## Core Helper Functions

### Text Frame Setup

```python
from pptx.enum.text import MSO_AUTO_SIZE

def setup_text_frame(tf, word_wrap=True):
    """Apply standard Cathay text frame settings.
    - 0.2cm margins all sides
    - No auto-shrink (MSO_AUTO_SIZE.NONE)
    """
    tf.word_wrap = word_wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE  # NEVER shrink to fit
    tf.margin_left = MARGIN_ALL    # 0.2cm
    tf.margin_right = MARGIN_ALL   # 0.2cm
    tf.margin_top = MARGIN_ALL     # 0.2cm
    tf.margin_bottom = MARGIN_ALL  # 0.2cm
```

### Paragraph Formatting (spacing + indent + line spacing)

```python
def format_paragraph(para, indent_left=True, is_bullet=False):
    """Apply standard Cathay paragraph formatting.
    - Spacing before: 4pt, after: 0pt
    - Line spacing: 1.1 (110%)
    - Left indent: 0.5cm (for bulleted content)
    """
    pPr = para._p.get_or_add_pPr()

    # Spacing before 4pt, after 0pt
    spcBef = pPr.find(qn('a:spcBef'))
    if spcBef is None:
        spcBef = etree.SubElement(pPr, qn('a:spcBef'))
    else:
        spcBef.clear()
    spcPts_bef = etree.SubElement(spcBef, qn('a:spcPts'))
    spcPts_bef.set('val', '400')  # 4pt = 400 hundredths

    spcAft = pPr.find(qn('a:spcAft'))
    if spcAft is None:
        spcAft = etree.SubElement(pPr, qn('a:spcAft'))
    else:
        spcAft.clear()
    spcPts_aft = etree.SubElement(spcAft, qn('a:spcPts'))
    spcPts_aft.set('val', '0')  # 0pt

    # Line spacing 1.1 (110%)
    lnSpc = pPr.find(qn('a:lnSpc'))
    if lnSpc is None:
        lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
    else:
        lnSpc.clear()
    spcPct = etree.SubElement(lnSpc, qn('a:spcPct'))
    spcPct.set('val', str(LINE_SPACING_PCT))  # 110000 = 1.1

    # Left indent 0.5cm for bulleted paragraphs
    if indent_left and is_bullet:
        pPr.set('indent', str(-Mm(3)))       # hanging indent for bullet
        pPr.set('marL', str(INDENT_LEFT))     # 0.5cm left margin
```

### Font Helper (auto Chinese/English detection)

```python
def set_run_font(run, text, size_pt=None, bold=False, color_rgb=None):
    """Set font with auto Chinese/English detection. Default size: 10pt."""
    size_pt = size_pt or DEFAULT_FONT_SIZE
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color_rgb:
        run.font.color.rgb = color_rgb

    has_chinese = any('\u4e00' <= c <= '\u9fff' for c in text)
    if has_chinese:
        run.font.name = "KaiTi"
        rPr = run._r.get_or_add_rPr()
        rPr.set(qn('a:altLang'), 'zh-CN')
        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = etree.SubElement(rPr, qn('a:ea'))
        ea.set('typeface', 'KaiTi')
    else:
        run.font.name = "Calibri"
```

### Filled Square Bullets (PPT XML buChar, 70%)

**CRITICAL: Use PPT's native bullet system (buChar XML), NOT text characters like "■ text".**

```python
def set_square_bullet(para, color='000000'):
    """Set filled square bullet via PPT buChar XML at 70% size.
    This uses PowerPoint's native bullet feature, not a text character.
    """
    pPr = para._p.get_or_add_pPr()
    # Remove any existing bullet settings
    for tag in ('a:buNone', 'a:buChar', 'a:buSzPct', 'a:buClr', 'a:buAutoNum', 'a:buFont'):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    # Bullet font (use Calibri for consistent square rendering)
    buFont = etree.SubElement(pPr, qn('a:buFont'))
    buFont.set('typeface', 'Calibri')
    # Bullet size: 70% of text size
    buSzPct = etree.SubElement(pPr, qn('a:buSzPct'))
    buSzPct.set('val', '70000')
    # Bullet color
    buClr = etree.SubElement(pPr, qn('a:buClr'))
    srgb = etree.SubElement(buClr, qn('a:srgbClr'))
    srgb.set('val', color)
    # Filled square character via buChar (PPT native bullet)
    buChar = etree.SubElement(pPr, qn('a:buChar'))
    buChar.set('char', '\u25A0')  # ■ BLACK SQUARE
```

### Add Bulleted Content

```python
def add_bullet_content(tf, items, size_pt=None, color_rgb=None):
    """Add bulleted content to a text frame.
    items: list of (text, level) tuples.
      level 0 = section header (bold red, no bullet, no indent)
      level 1+ = bulleted item (filled square, 0.5cm indent)
    """
    size_pt = size_pt or DEFAULT_FONT_SIZE
    color = color_rgb or CATHAY_BLACK
    for i, (text, level) in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.level = level
        p.alignment = PP_ALIGN.LEFT
        format_paragraph(p, indent_left=True, is_bullet=(level >= 1))

        run = p.add_run()

        if level == 0:
            set_run_font(run, text, size_pt=size_pt + 2, bold=True, color_rgb=CATHAY_RED)
        elif level == 1:
            set_run_font(run, text, size_pt=size_pt, color_rgb=color)
            set_square_bullet(p)
        elif level >= 2:
            set_run_font(run, text, size_pt=max(size_pt - 1, 8), color_rgb=RGBColor(0x80, 0x80, 0x80))
            set_square_bullet(p, color='808080')
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', str(Mm(10)))
            pPr.set('indent', str(-Mm(3)))
```

### Source Footer (8pt, bottom-left)

```python
def add_source_footer(slide, source_text):
    """Add 8pt source line at fixed position: left 1cm, top 18cm."""
    txBox = slide.shapes.add_textbox(
        Mm(10), Mm(180),    # left=1cm, top=18cm
        Mm(230), Mm(8))     # width ~23cm, height ~0.8cm
    tf = txBox.text_frame
    setup_text_frame(tf)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    format_paragraph(p, indent_left=False)
    run = p.add_run()
    set_run_font(run, f"Source: {source_text}", size_pt=8, color_rgb=CATHAY_GREY)
```

## Creating Presentations

```python
import os
prs = Presentation(TEMPLATE)

# DELETE all existing slides first
while len(prs.slides) > 0:
    rId = prs.slides._sldIdLst[0].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[0]
```

## Layout Reference

### Layout [0]: Red Title — Cover Slide

No title placeholder. All content via text boxes. **Left-aligned.**

```python
cover = prs.slides.add_slide(prs.slide_layouts[0])

# Fund name (top-left)
txBox = cover.shapes.add_textbox(
    Inches(CONTENT_LEFT), Inches(2.0), Inches(8.0), Inches(0.6))
tf = txBox.text_frame
setup_text_frame(tf)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
run = p.add_run()
set_run_font(run, "Cathay Smart Energy Fund", size_pt=20, color_rgb=CATHAY_RED)

# Company name + title (left-aligned, large)
txBox2 = cover.shapes.add_textbox(
    Inches(CONTENT_LEFT), Inches(3.0), Inches(8.5), Inches(1.2))
tf2 = txBox2.text_frame
setup_text_frame(tf2)
p2 = tf2.paragraphs[0]
p2.alignment = PP_ALIGN.LEFT
set_line_spacing(p2, 1.2)
run2 = p2.add_run()
set_run_font(run2, "Company Name", size_pt=28, bold=True, color_rgb=CATHAY_BLACK)
# Subtitle line
p2b = tf2.add_paragraph()
p2b.alignment = PP_ALIGN.LEFT
set_line_spacing(p2b, 1.2)
run2b = p2b.add_run()
set_run_font(run2b, "Investment Memo", size_pt=18, color_rgb=CATHAY_GREY)

# Date (left-aligned)
txBox3 = cover.shapes.add_textbox(
    Inches(CONTENT_LEFT), Inches(5.0), Inches(3.0), Inches(0.4))
tf3 = txBox3.text_frame
p3 = tf3.paragraphs[0]
p3.alignment = PP_ALIGN.LEFT
run3 = p3.add_run()
set_run_font(run3, "March 2026", size_pt=14, color_rgb=CATHAY_GREY)
```

### Layout [4]: 5_Red Slide — Main Content (most used)

| idx | Type | Position | Use |
|-----|------|----------|-----|
| 0 | TITLE | x=0.42", y=0.11", w=9.25", h=0.68" | Slide title (white on red bar) |
| 14 | DATE | x=0.47", y=7.15" | Date |
| 16 | SLIDE_NUMBER | x=7.28", y=7.15" | Page number |
| 3 | FOOTER | x=3.31", y=7.15" | Footer |

**Title is white text on dark red bar — set automatically by template.**

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])

# Set title (white on red bar — template handles color)
for shape in slide.shapes:
    if hasattr(shape, 'placeholder_format') and shape.placeholder_format is not None:
        if shape.placeholder_format.type == 1:  # TITLE
            shape.text = "Investment Thesis"

# Add body content with bullets
txBox = slide.shapes.add_textbox(
    Inches(CONTENT_LEFT), Inches(CONTENT_TOP),
    Inches(CONTENT_WIDTH), Inches(CONTENT_HEIGHT - 0.3))  # leave room for source
tf = txBox.text_frame
setup_text_frame(tf)

add_bullet_content(tf, [
    ("Market Opportunity", 0),           # section header (bold red, no bullet)
    ("30% CAGR in distributed solar", 1),  # ■ bullet
    ("C&I is most attractive segment", 1),
    ("Financial Highlights", 0),
    ("Revenue CAGR 45% (2023-2026E)", 1),
    ("Target equity IRR 15-20%", 1),
])

# Source footer (8pt, bottom-left)
add_source_footer(slide, "Company filings, Cathay analysis")
```

### Layout [11]: 1_Red Slide — Section Divider

Dark red background, white centered title. Used between major sections.

| idx | Type | Position | Use |
|-----|------|----------|-----|
| 0 | TITLE | x=0.87", y=4.19", w=8.32", h=0.68" | Section title (white, left-aligned) |

```python
divider = prs.slides.add_slide(prs.slide_layouts[11])
for shape in divider.shapes:
    if hasattr(shape, 'placeholder_format') and shape.placeholder_format is not None:
        if shape.placeholder_format.type == 1:
            shape.text = "1. Market Background"
            # White text on dark background
            for para in shape.text_frame.paragraphs:
                para.alignment = PP_ALIGN.LEFT
                for run in para.runs:
                    run.font.name = "Calibri"
                    run.font.color.rgb = CATHAY_WHITE
```

### Layout [5]: Red Slide — Content with Left Body

Title + left-half body placeholder. Use for text-heavy slides or slides with left text + right chart.

| idx | Type | Position | Use |
|-----|------|----------|-----|
| 0 | TITLE | x=0.42", y=0.11", w=9.25", h=0.68" | Title |
| 13 | BODY | x=0.47", y=1.52", w=4.33", h=2.28" | Left body text |

### Layout [6]: 3_Red Slide — Content + Bottom Notes

Title + left body + bottom notes area. Use when source/methodology notes are extensive.

| idx | Type | Position | Use |
|-----|------|----------|-----|
| 0 | TITLE | x=0.42", y=0.11" | Title |
| 13 | BODY | x=0.47", y=1.52", w=4.33" | Left body |
| 17 | BODY | x=0.47", y=6.17", w=9.06" | Bottom source/notes |

### Layout [8]: 2_Red Slide — Title Only (Blank Content)

Title bar only, rest is blank. Use for full-page charts, diagrams, or custom layouts.

| idx | Type | Position | Use |
|-----|------|----------|-----|
| 0 | TITLE | x=0.42", y=0.11" | Title |

### Layout [9]: Vide — Completely Blank

No placeholders. Use for full-bleed images or completely custom slides.

## PE Layout System

All layouts use content area: **left=1cm, top=2.92cm, width=23.4cm, height~14.6cm**.
Gap between columns/rows: **5mm horizontal, 3mm vertical**.

### Layout Grid Constants

```python
# Content area (mm)
CL = 10       # content left (1cm)
CT = 29.2     # content top (2.92cm)
CW = 234      # content width
CH = 146      # content height (to ~17.5cm, above source at 18cm)
GAP_H = 5     # horizontal gap between columns (mm)
GAP_V = 3     # vertical gap between rows (mm)

# Column widths (mm)
FULL   = CW                         # 234mm
HALF   = (CW - GAP_H) / 2          # 114.5mm
THIRD  = (CW - GAP_H * 2) / 3      # 74.7mm
QUARTER = (CW - GAP_H * 3) / 4     # 54.75mm
ONE_THIRD  = (CW - GAP_H) * 1/3    # 76.3mm
TWO_THIRDS = (CW - GAP_H) * 2/3    # 152.7mm
ONE_QUARTER  = (CW - GAP_H * 1) * 1/4  # 57.25mm
THREE_QUARTER = (CW - GAP_H * 1) * 3/4 # 171.75mm

# Row heights (mm)
ROW_FULL  = CH                       # 146mm
ROW_HALF  = (CH - GAP_V) / 2        # 71.5mm
ROW_THIRD = (CH - GAP_V * 2) / 3    # 46.7mm

# Column X positions (mm from left edge)
X1 = CL                              # 10mm — first column start
X2_HALF = CL + HALF + GAP_H          # 129.5mm — second half
X2_Q34 = CL + ONE_QUARTER + GAP_H    # 72.25mm — 3/4 start (after 1/4)
X2_T23 = CL + ONE_THIRD + GAP_H      # 91.3mm — 2/3 start (after 1/3)
X2_MID = CL + THIRD + GAP_H          # 89.7mm — second 1/3
X3_RIGHT = CL + THIRD*2 + GAP_H*2    # 169.4mm — third 1/3

# Row Y positions (mm from top edge)
Y1 = CT                               # 29.2mm — first row
Y2_HALF = CT + ROW_HALF + GAP_V       # 103.7mm — second half
Y2_MID = CT + ROW_THIRD + GAP_V       # 78.9mm — second 1/3 row
Y3_BOT = CT + ROW_THIRD*2 + GAP_V*2   # 128.6mm — third 1/3 row
```

### 1/2 + 1/2 (Equal Split)

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
# Left half
tL = slide.shapes.add_textbox(Mm(X1), Mm(CT), Mm(HALF), Mm(CH))
# Right half
tR = slide.shapes.add_textbox(Mm(X2_HALF), Mm(CT), Mm(HALF), Mm(CH))
```

### 1/4 + 3/4 (Sidebar + Main)

Use for: key stats sidebar + detailed content, or navigation + body.

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
# Left 1/4 (sidebar — often dark bg with white text for key metrics)
tL = slide.shapes.add_textbox(Mm(X1), Mm(CT), Mm(ONE_QUARTER), Mm(CH))
# Right 3/4 (main content)
tR = slide.shapes.add_textbox(Mm(X2_Q34), Mm(CT), Mm(THREE_QUARTER), Mm(CH))
```

### 1/3 + 2/3 (Narrow + Wide)

Use for: summary column + detail table, or chart + commentary.

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
# Left 1/3
tL = slide.shapes.add_textbox(Mm(X1), Mm(CT), Mm(ONE_THIRD), Mm(CH))
# Right 2/3
tR = slide.shapes.add_textbox(Mm(X2_T23), Mm(CT), Mm(TWO_THIRDS), Mm(CH))
```

### 2/3 + 1/3 (Wide + Narrow)

Use for: main content + sidebar summary, or chart + key takeaways.

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
# Left 2/3
tL = slide.shapes.add_textbox(Mm(X1), Mm(CT), Mm(TWO_THIRDS), Mm(CH))
# Right 1/3
tR = slide.shapes.add_textbox(Mm(X2_T23), Mm(CT), Mm(ONE_THIRD), Mm(CH))
```

### 3/4 + 1/4 (Main + Sidebar)

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
# Left 3/4
tL = slide.shapes.add_textbox(Mm(X1), Mm(CT), Mm(THREE_QUARTER), Mm(CH))
# Right 1/4
tR = slide.shapes.add_textbox(Mm(X2_Q34 + THREE_QUARTER - ONE_QUARTER), Mm(CT), Mm(ONE_QUARTER), Mm(CH))
```

### 1/3 + 1/3 + 1/3 (Three Columns)

Use for: comparing 3 scenarios, 3 business segments, 3 investment criteria.

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
t1 = slide.shapes.add_textbox(Mm(X1), Mm(CT), Mm(THIRD), Mm(CH))
t2 = slide.shapes.add_textbox(Mm(X2_MID), Mm(CT), Mm(THIRD), Mm(CH))
t3 = slide.shapes.add_textbox(Mm(X3_RIGHT), Mm(CT), Mm(THIRD), Mm(CH))
```

### Horizontal 1/3 + 1/3 + 1/3 (Three Rows)

Use for: timeline (past/present/future), stacked analysis sections, process flow.

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
# Top row
tTop = slide.shapes.add_textbox(Mm(X1), Mm(Y1), Mm(CW), Mm(ROW_THIRD))
# Middle row
tMid = slide.shapes.add_textbox(Mm(X1), Mm(Y2_MID), Mm(CW), Mm(ROW_THIRD))
# Bottom row
tBot = slide.shapes.add_textbox(Mm(X1), Mm(Y3_BOT), Mm(CW), Mm(ROW_THIRD))
```

### Four Quadrants (2x2 Grid)

Use for: SWOT, risk matrix, business overview (4 segments).

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
tTL = slide.shapes.add_textbox(Mm(X1), Mm(Y1), Mm(HALF), Mm(ROW_HALF))
tTR = slide.shapes.add_textbox(Mm(X2_HALF), Mm(Y1), Mm(HALF), Mm(ROW_HALF))
tBL = slide.shapes.add_textbox(Mm(X1), Mm(Y2_HALF), Mm(HALF), Mm(ROW_HALF))
tBR = slide.shapes.add_textbox(Mm(X2_HALF), Mm(Y2_HALF), Mm(HALF), Mm(ROW_HALF))
```

### KPI Metrics Row + Body

Use for: key figures header with supporting analysis below.

```python
from pptx.enum.shapes import MSO_SHAPE

def add_kpi_row(slide, kpis, y_mm=None):
    """kpis: list of (value, label) tuples."""
    y_mm = y_mm or CT
    n = len(kpis)
    bw = (CW - GAP_H * (n - 1)) / n
    for i, (val, lbl) in enumerate(kpis):
        x = X1 + i * (bw + GAP_H)
        sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Mm(x), Mm(y_mm), Mm(bw), Mm(22))
        sh.fill.solid(); sh.fill.fore_color.rgb = CATHAY_RED; sh.line.fill.background()
        tf = sh.text_frame; setup_text_frame(tf)
        pv = tf.paragraphs[0]; pv.alignment = PP_ALIGN.CENTER
        rv = pv.add_run(); set_run_font(rv, val, size_pt=16, bold=True, color_rgb=CATHAY_WHITE)
        pl = tf.add_paragraph(); pl.alignment = PP_ALIGN.CENTER
        rl = pl.add_run(); set_run_font(rl, lbl, size_pt=8, color_rgb=CATHAY_WHITE)
    return y_mm + 28  # return Y for content below KPIs

# Usage:
next_y = add_kpi_row(slide, [("$1.2bn", "Revenue"), ("45%", "CAGR"), ("15-20%", "IRR")])
body_box = slide.shapes.add_textbox(Mm(X1), Mm(next_y), Mm(CW), Mm(CT + CH - next_y))
```

### 1/4 Dark Sidebar + 3/4 Content (PE Deal Summary Style)

Classic PE memo layout: dark sidebar with key metrics, main area with details.

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])

# Dark sidebar (1/4)
sidebar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(X1), Mm(CT), Mm(ONE_QUARTER), Mm(CH))
sidebar.fill.solid(); sidebar.fill.fore_color.rgb = CATHAY_RED; sidebar.line.fill.background()
# Overlay text on sidebar
tS = slide.shapes.add_textbox(Mm(X1+2), Mm(CT+2), Mm(ONE_QUARTER-4), Mm(CH-4))
tfS = tS.text_frame; setup_text_frame(tfS)
add_bullet_content(tfS, [
    ("Key Metrics", 0),  # will be bold — override color to white
    ("EV: $500M", 1),
    ("Rev: $120M", 1),
    ("EBITDA: $35M", 1),
], color_rgb=CATHAY_WHITE)

# Main content (3/4)
tM = slide.shapes.add_textbox(Mm(X2_Q34), Mm(CT), Mm(THREE_QUARTER), Mm(CH))
tfM = tM.text_frame; setup_text_frame(tfM)
add_bullet_content(tfM, [("Company Overview", 0), ("Details here...", 1)])
```

### Top KPI + Bottom Three-Column Comparison

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
# KPI row at top
add_kpi_row(slide, [("Base", "Scenario"), ("Upside", "Scenario"), ("Downside", "Scenario")])
# Three columns below
col_y = CT + 28
col_h = CH - 28
t1 = slide.shapes.add_textbox(Mm(X1), Mm(col_y), Mm(THIRD), Mm(col_h))
t2 = slide.shapes.add_textbox(Mm(X2_MID), Mm(col_y), Mm(THIRD), Mm(col_h))
t3 = slide.shapes.add_textbox(Mm(X3_RIGHT), Mm(col_y), Mm(THIRD), Mm(col_h))
```

## PE Layout Selection Guide

| Slide Type | Recommended Layout | Example |
|-----------|-------------------|---------|
| Executive Summary | KPI row + full body | Key metrics + bullet points |
| Deal Overview | 1/4 dark sidebar + 3/4 | Key stats + company description |
| Investment Thesis | Full width bullets | Section headers + bullets |
| Risk/Mitigation | Full width table | 3-column risk table |
| Market Analysis | 1/3 + 2/3 | Key takeaways + chart/data |
| Business Model | 1/2 + 1/2 | Two segments side by side |
| Scenario Analysis | 1/3 + 1/3 + 1/3 columns | Base / Upside / Downside |
| Financials (P&L/BS/CF) | Full width table | Financial statement table |
| Valuation | 2/3 + 1/3 | Comps table + key multiples |
| Team / Org | Three rows (1/3 each) | Bios stacked vertically |
| Comps | Full width table | Trading comps |
| Returns | KPI + 1/2 + 1/2 | IRR/MOIC + sensitivity tables |
| Cap Table | Full width table | Ownership breakdown |
| Section Divider | Layout [11] | White text on dark red |

# Right column
right_box = slide.shapes.add_textbox(
    Inches(5.1), Inches(CONTENT_TOP),
    Inches(4.3), Inches(CONTENT_HEIGHT - 0.3))
tf_right = right_box.text_frame
setup_text_frame(tf_right)
add_bullet_content(tf_right, [
    ("Right Column Header", 0),
    ("Point A", 1),
    ("Point B", 1),
])
```

### Four-Quadrant Layout

```python
slide = prs.slides.add_slide(prs.slide_layouts[4])
# Set title...

quads = [
    (CONTENT_LEFT, CONTENT_TOP, 4.3, 2.8, "Q1 Title", ["Item 1", "Item 2"]),
    (5.1, CONTENT_TOP, 4.3, 2.8, "Q2 Title", ["Item A", "Item B"]),
    (CONTENT_LEFT, 3.9, 4.3, 2.8, "Q3 Title", ["Item X", "Item Y"]),
    (5.1, 3.9, 4.3, 2.8, "Q4 Title", ["Item P", "Item Q"]),
]
for x, y, w, h, title, items in quads:
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    setup_text_frame(tf)
    content = [(title, 0)] + [(item, 1) for item in items]
    add_bullet_content(tf, content)
```

### Key Metrics / KPI Row

```python
def add_kpi_row(slide, kpis, y=2.0):
    """Add a row of KPI callout boxes.
    kpis: list of (value, label) tuples, e.g. [("$500M", "Revenue"), ("45%", "CAGR")]
    """
    n = len(kpis)
    box_width = min(2.0, (CONTENT_WIDTH - 0.2 * (n - 1)) / n)
    gap = (CONTENT_WIDTH - box_width * n) / max(n - 1, 1)

    for i, (value, label) in enumerate(kpis):
        x = CONTENT_LEFT + i * (box_width + gap)
        shape = slide.shapes.add_shape(
            1,  # ROUNDED_RECTANGLE
            Inches(x), Inches(y), Inches(box_width), Inches(0.8))
        shape.fill.solid()
        shape.fill.fore_color.rgb = CATHAY_RED
        shape.line.fill.background()

        tf = shape.text_frame
        tf.word_wrap = True
        setup_text_frame(tf)
        # Value
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        set_run_font(run, value, size_pt=18, bold=True, color_rgb=CATHAY_WHITE)
        # Label
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run2 = p2.add_run()
        set_run_font(run2, label, size_pt=9, color_rgb=CATHAY_WHITE)
```

### Risk Table (Red/Amber/Green)

```python
def add_risk_table(slide, risks, top=1.2):
    """Add a risk/mitigation table.
    risks: list of (description, risk_text, mitigation_text) tuples.
    """
    headers = ["Description", "Risks", "Mitigations"]
    data = [headers] + [[r[0], r[1], r[2]] for r in risks]
    add_table(slide, len(data), 3, data, left=CONTENT_LEFT, top=top, width=CONTENT_WIDTH)
```

## Adding Tables

```python
def add_table(slide, rows, cols, data, left=None, top=1.2, width=None, row_height=0.35):
    """Add a Cathay-formatted table."""
    left = left or CONTENT_LEFT
    width = width or CONTENT_WIDTH
    table_shape = slide.shapes.add_table(
        rows, cols,
        Inches(left), Inches(top),
        Inches(width), Inches(rows * row_height))
    table = table_shape.table

    for i in range(rows):
        for j in range(cols):
            cell = table.cell(i, j)
            cell.text = str(data[i][j])
            # Set margins on cell text frame
            cell.margin_left = MARGIN_LEFT
            cell.margin_right = MARGIN_RIGHT
            cell.margin_top = Mm(3)
            cell.margin_bottom = Mm(3)

            for para in cell.text_frame.paragraphs:
                para.font.name = "Calibri"
                para.font.size = Pt(10)
                set_line_spacing(para, 1.2)

                if i == 0:
                    # Header row: white text, dark red bg
                    para.font.bold = True
                    para.font.color.rgb = CATHAY_WHITE
                else:
                    para.font.color.rgb = CATHAY_BLACK

            if i == 0:
                # Dark red header background
                tcPr = cell._tc.get_or_add_tcPr()
                solidFill = tcPr.makeelement(qn('a:solidFill'), {})
                srgbClr = solidFill.makeelement(qn('a:srgbClr'), {'val': '800000'})
                solidFill.append(srgbClr)
                tcPr.append(solidFill)
            elif i % 2 == 0:
                # Alternating light grey rows
                tcPr = cell._tc.get_or_add_tcPr()
                solidFill = tcPr.makeelement(qn('a:solidFill'), {})
                srgbClr = solidFill.makeelement(qn('a:srgbClr'), {'val': 'F2F2F2'})
                solidFill.append(srgbClr)
                tcPr.append(solidFill)

    return table
```

## Standard IC Memo Structure

| Section | Layout | Typical Slides |
|---------|--------|---------------|
| Cover | Red Title [0] | 1 |
| Overview / Status / Next Steps | 5_Red Slide [4] | 1 |
| Investment Thesis | 5_Red Slide [4] | 1 |
| Risks and Mitigations | 5_Red Slide [4] | 2 |
| Key Terms Summary | 5_Red Slide [4] | 2 |
| **Section: FDD & LDD Summary** | 1_Red Slide [11] | 1 |
| FDD/LDD Findings | 5_Red Slide [4] | 4-5 |
| **Section: Market Background** | 1_Red Slide [11] | 1 |
| Market Analysis | 5_Red Slide [4] / Red Slide [5] | 8-12 |
| **Section: Company Overview** | 1_Red Slide [11] | 1 |
| Company Details | Red Slide [5] / 5_Red Slide [4] | 6-8 |
| **Section: Financial Forecast & Return** | 1_Red Slide [11] | 1 |
| P&L, BS, CF, Valuation, Returns | 5_Red Slide [4] | 8-10 |
| **Section: Appendix** | 1_Red Slide [11] | 1 |
| Supplementary Material | Red Slide [5] | 2-3 |

## Excel → PPT Data Bridge

Read data from .xlsx models and populate PPT slides automatically.

```python
import openpyxl

def read_excel_table(xlsx_path, sheet_name, min_row, max_row, min_col, max_col):
    """Read a range from Excel and return as 2D list of strings."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name]
    data = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        data.append([str(cell.value) if cell.value is not None else "" for cell in row])
    wb.close()
    return data


def excel_to_ppt_table(slide, xlsx_path, sheet_name, min_row, max_row, min_col, max_col,
                        left_mm=None, top_mm=None, width_mm=None):
    """Read Excel range → generate Cathay-styled PPT table on slide."""
    left_mm = left_mm or CONTENT_LEFT_CM * 10
    top_mm = top_mm or CONTENT_TOP_CM * 10
    width_mm = width_mm or 234
    data = read_excel_table(xlsx_path, sheet_name, min_row, max_row, min_col, max_col)
    rows, cols = len(data), len(data[0]) if data else 0
    # Calls add_table from the Tables section above
    return add_table(slide, rows, cols, data, left=left_mm, top=top_mm, width=width_mm)


def excel_to_kpi_row(slide, xlsx_path, cells_map, y_mm=None):
    """Read specific cells from Excel → generate KPI metric boxes.
    cells_map: list of (sheet, cell_ref, label) e.g. [("Summary", "B5", "Revenue"), ...]
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    kpis = []
    for sheet, cell, label in cells_map:
        val = wb[sheet][cell].value
        kpis.append((str(val) if val else "N/A", label))
    wb.close()
    return add_kpi_row(slide, kpis, y_mm=y_mm)
```

**Usage:**
```python
# Read comps table from Excel and put into PPT
excel_to_ppt_table(slide, "AAPL_Comps_20260317.xlsx", "Comps", 1, 15, 1, 8)

# Read key metrics for KPI row
excel_to_kpi_row(slide, "Model.xlsx", [
    ("Summary", "B3", "Revenue"),
    ("Summary", "B7", "EBITDA"),
    ("Returns", "B12", "IRR"),
    ("Returns", "B13", "MOIC"),
])
```

## Chart Generation (Cathay Brand)

Generate matplotlib charts with Cathay brand palette, export as PNG, insert into PPT.

```python
import matplotlib
matplotlib.use('Agg')  # non-interactive backend
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import tempfile, os

# Cathay brand color cycle for charts
CATHAY_COLORS = ['#800000', '#C8A415', '#808080', '#E60000', '#E8D590', '#404040', '#D9D9D9']

def _setup_cathay_style():
    """Apply Cathay brand styling to matplotlib."""
    # Use Arial as fallback if Calibri not available in matplotlib
    import matplotlib.font_manager as fm
    available = {f.name for f in fm.fontManager.ttflist}
    font = 'Calibri' if 'Calibri' in available else 'Arial'
    plt.rcParams.update({
        'font.family': font,
        'font.size': 10,
        'axes.prop_cycle': plt.cycler(color=CATHAY_COLORS),
        'axes.edgecolor': '#808080',
        'axes.linewidth': 0.5,
        'grid.color': '#D9D9D9',
        'grid.linewidth': 0.5,
        'figure.facecolor': 'white',
        'axes.facecolor': 'white',
    })


def cathay_bar_chart(categories, values, title, output_path=None, ylabel="", figsize=(8, 4.5)):
    """Cathay brand bar chart. Returns path to saved PNG."""
    _setup_cathay_style()
    fig, ax = plt.subplots(figsize=figsize)
    bars = ax.bar(categories, values, color=CATHAY_COLORS[:len(values)], width=0.6)
    ax.set_title(title, fontsize=12, fontweight='bold', color='#800000', loc='left')
    ax.set_ylabel(ylabel, fontsize=9, color='#808080')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', alpha=0.3)
    # Value labels on bars
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(values)*0.02,
                f'{val:,.0f}' if isinstance(val, (int, float)) else str(val),
                ha='center', va='bottom', fontsize=9, color='#404040')
    plt.tight_layout()
    output_path = output_path or os.path.join(tempfile.gettempdir(), 'cathay_bar.png')
    fig.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return output_path


def cathay_line_chart(x_labels, series_dict, title, output_path=None, ylabel="", figsize=(8, 4.5)):
    """Cathay brand line chart. series_dict: {"Series A": [vals], "Series B": [vals]}."""
    _setup_cathay_style()
    fig, ax = plt.subplots(figsize=figsize)
    for i, (name, vals) in enumerate(series_dict.items()):
        ax.plot(x_labels, vals, marker='o', markersize=4, linewidth=2,
                color=CATHAY_COLORS[i % len(CATHAY_COLORS)], label=name)
    ax.set_title(title, fontsize=12, fontweight='bold', color='#800000', loc='left')
    ax.set_ylabel(ylabel, fontsize=9, color='#808080')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.legend(fontsize=9, frameon=False)
    ax.grid(axis='y', alpha=0.3)
    plt.tight_layout()
    output_path = output_path or os.path.join(tempfile.gettempdir(), 'cathay_line.png')
    fig.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return output_path


def cathay_waterfall_chart(labels, values, title, output_path=None, figsize=(8, 4.5)):
    """Waterfall chart (common in PE for value bridge / EV walk)."""
    _setup_cathay_style()
    fig, ax = plt.subplots(figsize=figsize)
    cumulative = [0]
    for v in values[:-1]:
        cumulative.append(cumulative[-1] + v)
    colors = []
    for i, v in enumerate(values):
        if i == 0 or i == len(values) - 1:
            colors.append('#800000')  # total bars
        elif v >= 0:
            colors.append('#C8A415')  # positive
        else:
            colors.append('#E60000')  # negative
    bottoms = cumulative[:-1] + [0]  # last bar starts from 0
    ax.bar(labels, [abs(v) for v in values], bottom=[max(0, b) if i < len(values)-1 else 0
           for i, b in enumerate(bottoms)], color=colors, width=0.5)
    ax.set_title(title, fontsize=12, fontweight='bold', color='#800000', loc='left')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.xticks(rotation=30, ha='right', fontsize=9)
    plt.tight_layout()
    output_path = output_path or os.path.join(tempfile.gettempdir(), 'cathay_waterfall.png')
    fig.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return output_path


def cathay_pie_chart(labels, values, title, output_path=None, figsize=(6, 4.5)):
    """Cathay brand pie/donut chart."""
    _setup_cathay_style()
    fig, ax = plt.subplots(figsize=figsize)
    wedges, texts, autotexts = ax.pie(values, labels=labels, autopct='%1.0f%%',
        colors=CATHAY_COLORS[:len(values)], startangle=90,
        wedgeprops={'linewidth': 1, 'edgecolor': 'white'})
    for t in autotexts:
        t.set_fontsize(9); t.set_color('white'); t.set_fontweight('bold')
    for t in texts:
        t.set_fontsize(9)
    ax.set_title(title, fontsize=12, fontweight='bold', color='#800000', loc='left')
    plt.tight_layout()
    output_path = output_path or os.path.join(tempfile.gettempdir(), 'cathay_pie.png')
    fig.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return output_path


def insert_chart_image(slide, image_path, x_mm=None, y_mm=None, w_mm=None):
    """Insert a chart image (width-only, preserves aspect ratio).
    DEPRECATED: prefer safe_chart_insert() which returns bottom_y_mm.
    """
    x_mm = x_mm or CONTENT_LEFT_CM * 10
    y_mm = y_mm or CONTENT_TOP_CM * 10
    w_mm = w_mm or 200
    slide.shapes.add_picture(image_path, Mm(x_mm), Mm(y_mm), Mm(w_mm))
```

## Safe Layout Helpers (Overflow Prevention)

```python
from PIL import Image

def safe_chart_insert(slide, image_path, x_mm=None, y_mm=None, w_mm=200):
    """Insert chart PNG with width-only sizing, return actual bottom Y (mm).

    Reads actual PNG pixel dimensions, computes rendered height preserving
    aspect ratio. Auto-scales if chart would exceed content zone.

    Returns:
        bottom_y_mm (float): the Y coordinate where the chart ends.
        Use this + GAP_V as the top of the next element below.
    """
    x_mm = x_mm or CL
    y_mm = y_mm or CT

    with Image.open(image_path) as img:
        px_w, px_h = img.size
    aspect = px_h / px_w
    rendered_h_mm = w_mm * aspect

    bottom_y = y_mm + rendered_h_mm
    if bottom_y > CONTENT_BOTTOM_MM:
        max_h = CONTENT_BOTTOM_MM - y_mm
        w_mm = max_h / aspect
        rendered_h_mm = max_h
        bottom_y = y_mm + rendered_h_mm

    slide.shapes.add_picture(image_path, Mm(x_mm), Mm(y_mm), Mm(w_mm))
    return bottom_y


def safe_textbox(slide, x_mm, y_mm, w_mm, h_mm=None, max_bottom_mm=None):
    """Create a textbox that respects content zone bounds.

    If h_mm is None, fills from y_mm to max_bottom_mm (default 175mm).
    Clamps height to never exceed content zone. Returns (shape, text_frame).
    """
    max_bottom_mm = max_bottom_mm or CONTENT_BOTTOM_MM
    if h_mm is None:
        h_mm = max_bottom_mm - y_mm

    actual_bottom = y_mm + h_mm
    if actual_bottom > max_bottom_mm:
        h_mm = max_bottom_mm - y_mm

    if h_mm <= 0:
        h_mm = 10  # minimum fallback

    txBox = slide.shapes.add_textbox(Mm(x_mm), Mm(y_mm), Mm(w_mm), Mm(h_mm))
    tf = txBox.text_frame
    setup_text_frame(tf)
    return txBox, tf


CONTENT_BOTTOM_MM = 175  # usable content ends here (source footer at 180mm)
```

## PPT QC: Overlap Detection & PDF Export

```python
import subprocess

def validate_no_overlap(pptx_path):
    """Check all slides for overlapping shapes. Returns list of issues."""
    prs_check = Presentation(pptx_path)
    issues = []
    for slide_idx, slide in enumerate(prs_check.slides, 1):
        shapes = []
        for sh in slide.shapes:
            l = sh.left / 914400; t = sh.top / 914400
            r = l + sh.width / 914400; b = t + sh.height / 914400
            shapes.append((sh.name, l, t, r, b))
        for i in range(len(shapes)):
            for j in range(i + 1, len(shapes)):
                n1, l1, t1, r1, b1 = shapes[i]
                n2, l2, t2, r2, b2 = shapes[j]
                if l1 < r2 and r1 > l2 and t1 < b2 and b1 > t2:
                    # Check if one is fully inside the other (sidebar pattern — OK)
                    inside = (l2 >= l1 and r2 <= r1 and t2 >= t1 and b2 <= b1) or \
                             (l1 >= l2 and r1 <= r2 and t1 >= t2 and b1 <= b2)
                    if not inside:
                        issues.append(f"Slide {slide_idx}: '{n1}' overlaps '{n2}'")
    return issues


def validate_text_fit(pptx_path):
    """Estimate whether text content fits within each textbox. Returns warnings."""
    prs_check = Presentation(pptx_path)
    warnings = []
    for slide_idx, slide in enumerate(prs_check.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            tf = shape.text_frame
            total_text = "".join(p.text for p in tf.paragraphs)
            if not total_text.strip():
                continue
            box_w_mm = shape.width / 36000
            box_h_mm = shape.height / 36000
            has_chinese = any('\u4e00' <= c <= '\u9fff' for c in total_text)
            chars_per_mm = 2.5 if has_chinese else 3.5
            usable_w = box_w_mm - 4
            chars_per_line = max(usable_w * chars_per_mm, 1)
            num_paragraphs = len([p for p in tf.paragraphs if p.text.strip()])
            total_chars = len(total_text)
            est_lines = (total_chars / chars_per_line) + num_paragraphs * 0.3
            line_height_mm = 2.8
            est_height = est_lines * line_height_mm + 4
            if est_height > box_h_mm * 1.15:
                overflow_pct = ((est_height - box_h_mm) / box_h_mm) * 100
                warnings.append(
                    f"Slide {slide_idx}: '{shape.name}' text may overflow "
                    f"(est {est_height:.0f}mm vs box {box_h_mm:.0f}mm, +{overflow_pct:.0f}%)")
    return warnings


def export_to_pdf(pptx_path, output_dir=None):
    """Convert PPTX to PDF via LibreOffice for visual QC."""
    output_dir = output_dir or os.path.dirname(pptx_path)
    result = subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, pptx_path],
        capture_output=True, text=True, timeout=120)  # LibreOffice first launch can be slow
    pdf_path = os.path.join(output_dir, os.path.splitext(os.path.basename(pptx_path))[0] + '.pdf')
    return pdf_path if os.path.exists(pdf_path) else None


def qc_presentation(pptx_path):
    """Run full QC: overlap check + text fit check + PDF export."""
    issues = validate_no_overlap(pptx_path)
    fit_warnings = validate_text_fit(pptx_path)
    if issues:
        print(f"OVERLAP ISSUES ({len(issues)}):")
        for issue in issues:
            print(f"  - {issue}")
    if fit_warnings:
        print(f"TEXT FIT WARNINGS ({len(fit_warnings)}):")
        for w in fit_warnings:
            print(f"  - {w}")
    if not issues and not fit_warnings:
        print("No layout issues found.")
    pdf = export_to_pdf(pptx_path)
    if pdf:
        print(f"PDF exported: {pdf}")
    return issues + fit_warnings, pdf
```

## Section Icons (Visual Navigation Markers)

```python
# Icon type constants: (MSO shape name, color hex)
ICON_FINANCIAL = (MSO_SHAPE.ROUNDED_RECTANGLE, 'C8A415')  # gold square — financial data
ICON_INSIGHT   = (MSO_SHAPE.OVAL, '800000')                # red circle — key insight
ICON_RISK      = (MSO_SHAPE.ISOSCELES_TRIANGLE, 'E60000')  # red triangle — risk/warning
ICON_CATALYST  = (MSO_SHAPE.DIAMOND, 'C8A415')             # gold diamond — catalyst
ICON_ACTION    = (MSO_SHAPE.RIGHT_ARROW, '800000')         # red arrow — action item

def add_section_marker(slide, x_mm, y_mm, icon_type=None):
    """Place a small colored shape (4x4mm) as a section visual marker."""
    icon_type = icon_type or ICON_INSIGHT
    shape_enum, color_hex = icon_type
    marker = slide.shapes.add_shape(shape_enum, Mm(x_mm), Mm(y_mm), Mm(4), Mm(4))
    marker.fill.solid()
    marker.fill.fore_color.rgb = RGBColor.from_string(color_hex)
    marker.line.fill.background()
    return marker

_ICON_KEYWORD_MAP = {
    ICON_FINANCIAL: ["收入", "利润", "EPS", "现金流", "季度", "财务", "Revenue", "Margin", "Cash", "资产"],
    ICON_RISK: ["风险", "Bear", "威胁", "下行", "Risk", "Kill", "止损"],
    ICON_CATALYST: ["催化", "时间表", "Catalyst", "Trigger", "监控", "Monitoring"],
    ICON_ACTION: ["行动", "决策", "建议", "Action", "Plan", "CIO", "裁决"],
    ICON_INSIGHT: ["论点", "Thesis", "Bull", "洞察", "优势", "Moat", "Industry"],
}

def auto_assign_icons(items):
    """Auto-assign icons to level-0 items based on keyword matching."""
    icons = {}
    for text, level in items:
        if level != 0:
            continue
        for icon_type, keywords in _ICON_KEYWORD_MAP.items():
            if any(kw in text for kw in keywords):
                icons[text] = icon_type
                break
        else:
            icons[text] = ICON_INSIGHT
    return icons
```

## Conclusion-Style Slide Titles

```python
def set_title(slide, title_text, size_pt=20):
    """Set slide title (white on red bar) — single-part version."""
    for shape in slide.shapes:
        if hasattr(shape, 'placeholder_format') and shape.placeholder_format is not None:
            if shape.placeholder_format.type == 1:
                tf = shape.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT
                run = p.add_run()
                set_run_font(run, title_text, size_pt=size_pt, bold=True, color_rgb=RGBColor(0xFF, 0xFF, 0xFF))
                break


def set_title_with_conclusion(slide, topic, conclusion):
    """Set slide title: 'Topic — Conclusion'. Topic=white, Conclusion=gold.

    Use for all Layout [4] content slides. Section dividers use set_title().
    Example: set_title_with_conclusion(slide, "投资摘要", "BUY, 目标价$520 (+14%)")
    """
    CATHAY_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    CATHAY_GOLD_RGB = RGBColor(0xC8, 0xA4, 0x15)
    for shape in slide.shapes:
        if hasattr(shape, 'placeholder_format') and shape.placeholder_format is not None:
            if shape.placeholder_format.type == 1:
                tf = shape.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT
                # Topic — white
                run_topic = p.add_run()
                set_run_font(run_topic, topic, size_pt=20, bold=True, color_rgb=CATHAY_WHITE)
                # Separator — gold
                run_sep = p.add_run()
                set_run_font(run_sep, " — ", size_pt=20, bold=False, color_rgb=CATHAY_GOLD_RGB)
                # Conclusion — gold, slightly smaller
                run_conc = p.add_run()
                set_run_font(run_conc, conclusion, size_pt=18, bold=True, color_rgb=CATHAY_GOLD_RGB)
                break
```

## IC Memo (.docx) → Cathay PPT Conversion

Convert a Word IC memo (from `private-equity:ic-memo` skill) into Cathay-branded PPT deck.

```python
from docx import Document

# Section header → layout mapping
IC_SECTION_MAP = {
    "Executive Summary": (4, "full_bullets"),
    "Company Overview": (4, "full_bullets"),
    "Market Analysis": (4, "full_bullets"),
    "Market Overview": (4, "full_bullets"),
    "Financial Analysis": (4, "full_bullets"),
    "Financial Overview": (4, "full_bullets"),
    "Investment Thesis": (4, "full_bullets"),
    "Deal Terms": (4, "full_bullets"),
    "Key Terms": (4, "full_bullets"),
    "Returns Analysis": (4, "full_bullets"),
    "Risk Factors": (4, "full_bullets"),
    "Risks": (4, "full_bullets"),
    "Recommendation": (4, "full_bullets"),
}

# Major sections that get a divider slide
IC_DIVIDER_SECTIONS = {
    "Market Analysis", "Market Overview",
    "Company Overview",
    "Financial Analysis", "Financial Overview",
    "Returns Analysis",
    "Risk Factors", "Risks",
    "Recommendation",
}


def docx_to_cathay_ppt(docx_path, output_pptx_path, fund_name="Cathay Smart Energy Fund"):
    """Convert IC Memo .docx → Cathay branded .pptx."""
    doc = Document(docx_path)
    prs_out = Presentation(TEMPLATE)
    # Clear existing slides
    while len(prs_out.slides) > 0:
        rId = prs_out.slides._sldIdLst[0].rId
        prs_out.part.drop_rel(rId)
        del prs_out.slides._sldIdLst[0]

    # Extract sections from docx
    sections = []
    current_section = None
    current_content = []
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading 1') or para.style.name.startswith('Heading 2'):
            if current_section:
                sections.append((current_section, current_content))
            current_section = para.text.strip()
            current_content = []
        elif para.text.strip():
            current_content.append(para.text.strip())
    if current_section:
        sections.append((current_section, current_content))

    # Generate cover slide
    cover = prs_out.slides.add_slide(prs_out.slide_layouts[0])
    # ... (use standard cover pattern from above)

    # Generate content slides
    section_num = 0
    for title, content in sections:
        # Add divider if major section
        if title in IC_DIVIDER_SECTIONS:
            section_num += 1
            div = prs_out.slides.add_slide(prs_out.slide_layouts[11])
            set_title(div, f"{section_num}. {title}")

        # Add content slide
        slide = prs_out.slides.add_slide(prs_out.slide_layouts[4])
        set_title(slide, title)
        tf = add_content_box(slide)
        items = [(line, 1) for line in content[:12]]  # max 12 bullets per slide
        add_bullet_content(tf, items)

        # Overflow to second slide if needed
        if len(content) > 12:
            slide2 = prs_out.slides.add_slide(prs_out.slide_layouts[4])
            set_title(slide2, f"{title} (cont.)")
            tf2 = add_content_box(slide2)
            add_bullet_content(tf2, [(line, 1) for line in content[12:24]])

    prs_out.save(output_pptx_path)
    return output_pptx_path
```

## Excel Brand Styling (Post-process)

Apply Cathay brand colors to any Excel model output.

```python
def apply_cathay_style(xlsx_path, output_path=None):
    """Post-process an Excel file to use Cathay brand colors.
    Replaces default blue headers (#1F4E79) with Cathay red (#800000).
    """
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.load_workbook(xlsx_path)
    cathay_fill = PatternFill(start_color='800000', end_color='800000', fill_type='solid')
    white_font = Font(name='Calibri', color='FFFFFF', bold=True)
    body_font = Font(name='Calibri', color='000000', size=10.5)

    for ws in wb.worksheets:
        # Restyle header row (row 1) if it has colored fills
        for cell in ws[1]:
            if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                rgb = str(cell.fill.start_color.rgb)
                if rgb not in ('00000000', 'FFFFFF', '00FFFFFF'):  # has a color
                    cell.fill = cathay_fill
                    cell.font = white_font
        # Set body font
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if cell.value is not None:
                    cell.font = body_font

    output_path = output_path or xlsx_path
    wb.save(output_path)
    return output_path
```

## Critical: No Overlap Rule

**NEVER allow text boxes, shapes, or tables to overlap each other.** Every element must have its own exclusive space on the slide. Before placing any element, verify:

1. **No X-axis overlap**: Element's `left + width` must not exceed the next element's `left`
2. **No Y-axis overlap**: Element's `top + height` must not exceed the next element's `top`
3. **Sidebar + overlay text**: When using dark sidebar pattern (shape + textbox), the textbox must be **inside** the shape bounds (inset by 2-3mm each side), not extending beyond it
4. **Source footer**: Always at fixed position `top=18cm` — content above must end before 17.5cm
5. **Title bar**: Template title placeholder occupies `y=0 to ~2.5cm` — content starts at `top=2.92cm`, never earlier
6. **Tables**: Table `top + (rows × row_height)` must not exceed 17.5cm (source area)

When computing layout grids, always account for gaps:
- Horizontal gap between columns: **5mm**
- Vertical gap between rows: **3mm**
- Column width = `(available_width - gaps) / n`, NOT `available_width / n`

**If content might overflow**, reduce font size or split across two slides. Never let text auto-shrink (`MSO_AUTO_SIZE.NONE` is enforced).

## Quality Checklist

Before finalizing any deck:
- [ ] All English/numbers use **Calibri**, all Chinese uses **楷体 (KaiTi)**
- [ ] Default body font size is **10.5pt**
- [ ] All titles and cover are **left-aligned**
- [ ] **White text** on all dark/red backgrounds
- [ ] All text boxes: margins **0.2cm** all sides, **no auto-shrink** (`MSO_AUTO_SIZE.NONE`)
- [ ] Paragraph spacing: before **4pt**, after **0pt**, line spacing **1.1**
- [ ] Bulleted paragraphs: left indent **0.5cm**
- [ ] Bullets use **PPT native buChar XML** (filled square at 70%), NOT text "■"
- [ ] Brand color `#800000` on title bars, table headers, callout boxes
- [ ] **8pt source footer** at bottom-left of every data slide
- [ ] Section dividers use Layout [11] with white text
- [ ] Page numbers present
- [ ] No text overlapping title bar or footer area
- [ ] `safe_chart_insert()` used for all chart insertions (no direct `add_picture` for charts)
- [ ] `safe_textbox()` used for all variable-content text boxes
- [ ] `validate_text_fit()` returns zero warnings
- [ ] All Layout [4] slides use `set_title_with_conclusion()` (not plain `set_title`)
- [ ] Content slides with 3+ topics use level-0 headers via `add_bullet_content()`
- [ ] Section markers (icons) present on slides with 2+ level-0 headers
- [ ] No direct `.font.name =` outside of `set_run_font()` function
- [ ] `save_with_validation()` used (includes `full_cleanup()` anti-corruption)
- [ ] All lines are thin rectangles, NEVER connectors
- [ ] `check_guard_rails()` returns zero violations
- [ ] 25+ slide decks use 5+ different grid patterns

## IRON RULE: Anti-Corruption Defense (v2, borrowed from McKinsey pattern)

**ALL lines must be thin rectangles, NEVER connectors.** Connectors carry corrupt `<p:style>` references that leak theme XML and corrupt PowerPoint rendering.

Three-layer defense:
1. **`_clean_shape(shape)`** — call after creating any shape; strips `<p:style>` XML inline
2. **`full_cleanup(pptx_path)`** — called automatically by `save_with_validation()`; rewrites the .pptx zip, regex-stripping ALL remaining `<p:style>` from XML files
3. **Connector ban** — use `slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Mm(0.2))` for lines

```python
# Drawing a horizontal separator line (CORRECT)
line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(CL), Mm(y), Mm(CW), Mm(0.3))
line.fill.solid()
line.fill.fore_color.rgb = CATHAY_LTGREY
line.line.fill.background()

# WRONG — never use connectors
# slide.shapes.add_connector(...)  # ← BANNED
```

## Accent Pairs (multi-item color cycling)

For layouts with 3+ parallel items (comparison columns, funnel stages, chart segments):

```python
from constants import CATHAY_ACCENT_PAIRS
# [(dark_red, light_red_bg), (gold, light_gold_bg), (accent_red, light_red_bg), (grey, light_grey_bg)]

for i, item in enumerate(items):
    fg_color, bg_color = CATHAY_ACCENT_PAIRS[i % len(CATHAY_ACCENT_PAIRS)]
    box.fill.solid()
    box.fill.fore_color.rgb = bg_color
    # header text uses fg_color
```

## 8 Production Guard Rails

Run `check_guard_rails(prs)` before save. Returns violation list.

| # | Rule | Threshold |
|---|------|-----------|
| 1 | Content bottom gap | Shapes must stop >=1mm before footer (180mm) |
| 2 | Right margin | All shapes right edge <= 244mm |
| 3 | Bottom whitespace | Lowest content within 8mm of footer zone |
| 4 | Horizontal overflow | `item_w * n + gap * (n-1) <= CW` |
| 5 | Peer font harmony | Same-Y shapes (±2mm) must share font size |
| 6 | Collision detection | Non-contained shapes >=0.8mm apart |
| 7 | Layout variety | 5+ unique grids per 25 slides; no 3 consecutive same |
| 8 | CJK density | Character count vs box-height density limits |

## AutoFix Pipeline (4-stage)

Run `autofix_pipeline(pptx_path)` for automatic fixes in priority order:

```
Stage 1: remove_redundancy  — deduplicate identical paragraphs
Stage 2: compress_text      — collapse multi-spaces, strip whitespace
Stage 3: restructure_layout — cap shapes at content zone, fix margins
Stage 4: font_micro_adjust  — reduce fonts (min: title≥18pt, body≥9pt, footer≥7pt)
```

## New Slide Templates (T11-T16)

| Template | Function | Use Case |
|----------|----------|----------|
| T11 | `template_donut_chart(prs, title, subtitle, segments, insight_bullets, source)` | Revenue split, market share — donut PNG + insight panel |
| T12 | `template_before_after(prs, title, subtitle, before_items, after_items, source)` | Process improvement, strategy shift — grey "Before" vs gold "After" |
| T13 | `template_funnel(prs, title, subtitle, stages, source)` | Sales pipeline, TAM/SAM/SOM — progressively narrowing bars |
| T14 | `template_swot(prs, title, subtitle, strengths, weaknesses, opportunities, threats, source)` | Color-coded 2x2 SWOT matrix |
| T15 | `template_waterfall(prs, title, subtitle, items, source)` | Value bridge, EV walk — matplotlib waterfall chart |
| T16 | `template_stakeholder_map(prs, title, subtitle, stakeholders, source)` | Relationship/stakeholder map — center + surrounding nodes |

## Incremental Insert Workflow (增量修改)

**IRON RULE: Once a user has manually edited a .pptx, NEVER regenerate the full deck.** Only insert/modify individual slides.

```python
# 1. Open user's edited version
prs = Presentation("user_draft.pptx")

# 2. Print current structure to find insertion point
for i, slide in enumerate(prs.slides, 1):
    title = ""
    for shape in slide.shapes:
        if hasattr(shape, 'placeholder_format') and shape.placeholder_format:
            if shape.placeholder_format.type == 1:
                title = shape.text_frame.text[:50]
    print(f"  Slide {i}: {title or '(no title)'}")

# 3. Add new slide (always appends to end)
red_layout = prs.slide_layouts[4]  # 5_Red Slide
new_slide = prs.slides.add_slide(red_layout)
clear_slide(new_slide)  # strip inherited placeholders

# 4. Build content on new slide
set_title_with_conclusion(new_slide, "新增分析", "关键发现")
# ... add content ...

# 5. Reorder to desired position (e.g., insert after slide 5)
n = len(prs.slides)
order = list(range(1, n)) + [n]  # default: current order
order.insert(5, order.pop())     # move last (new) slide to position 6
reorder_slides(prs, order)

# 6. Save as NEW file (never overwrite user's version)
prs.save("user_draft_v2.pptx")
```

Key functions:
- `reorder_slides(prs, [1, 2, 5, 3, 4])` — reorder after append
- `clear_slide(slide)` — strip all shapes for rebuild
- `add_multi_text(slide, x, y, w, h, segments)` — flexible multi-paragraph text

## Multi-Paragraph Text (`add_multi_text`)

More flexible than `add_bullet_content` — each paragraph has independent formatting:

```python
add_multi_text(slide, CL, CT, CW, 50, [
    ("核心观点", dict(size=14, bold=True, color=CATHAY_RED, space_after=4)),
    ("AI算力需求驱动数据中心建设加速", dict(size=10)),
    ("预计2030年市场规模达$430B (IEA)", dict(size=10, italic=True, color=CATHAY_GREY)),
], fill_rgb=RGBColor(0xF2, 0xF2, 0xF2))  # optional light grey background
```

## IC Memo Content Discipline

### Reading Order (动笔前必看)
1. **管访 .md** — founder/CTO latest numbers, valuation & milestones (most authoritative)
2. **DD / 投资报告 .docx** — financials, team table, legal structure
3. **研究简报 .pptx** — industry TAM, market sizing, competitor landscape
4. **Captable .xlsx** — ownership structure

### Number Rules
- **每数字必有出处** — if a number can't be traced to source material, delete it or use ranges
- Valuation: write "本轮投前 X 亿" not "估值 X 亿"
- CAPEX: provide two cross-checking metrics (e.g., "万元/kW" + total per unit)
- Market size: always include year + methodology + source (e.g., "2030 全球 AI IDC ~30 GW (IEA 2025)")
- Team: use the most recent interview numbers, not older DD data
- **Uncertain numbers → delete or use ranges** ("~Y 名工程师" not "38 名工程师")

### Tone
- IC Memo is objective presentation, NOT an approve/reject recommendation
- Balance every "highlight" with a "watch item"
- Conclusion: "建议投委会关注 X / Y / Z" — never "建议投" or "不建议投"
- Avoid absolutes: "唯一" → "少数几家之一"; "全球领先" → "位列第一梯队"

## Visual QA Sub-Agent Pattern

After generating any deck, always render to JPG and dispatch a fresh-eyes sub-agent:

```bash
soffice --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 140 output.pdf slide
```

Sub-agent prompt template:
```
检查 slide-N.jpg. 假设有问题 — 去找.
重点: 重叠 / 溢出 / 标题 wrap / 间距异常 / 对比度差 / 占位符残留.
列出所有问题.
```

**Why sub-agent**: You "know what the slide should look like" and your eyes auto-complete. A fresh agent sees raw pixels without that bias.

## Common Pitfalls (速查)

| Symptom | Root Cause | Fix |
|---------|-----------|-----|
| 中文变 Calibri 方块 | Missing `a:ea` + `a:cs` XML | Use `set_run_font()` (never `.font.name` directly) |
| 左侧红竖条消失 | Wrong slide layout | Use `prs.slide_layouts[4]` ("5_Red Slide") |
| Logo被遮挡 | Shape overlaps bottom-right | Keep shapes within `x+w ≤ 218mm`, `y+h ≤ 175mm` |
| Title和顶部横线打架 | Drew extra red bar on Layout [4] | Layout [4] already has chrome — don't redraw |
| 文本溢出底部 | Font too large / line spacing too wide | Reduce 0.5pt, or `line_spacing` 1.2→1.1, or split to 2 columns |
| 表格和下方元素贴脸 | No gap | Reduce table height by GAP_V, add spacing |
| 增量插入后页码错位 | python-pptx appends to end | Use `reorder_slides()` after append |

## Anti-Patterns (禁止)

1. **从零画红色 chrome** — Layout [4] already has the red bar, left stripe, and logo. Never redraw.
2. **prs.save() 覆盖用户手改** — Always save as `_v2.pptx`. Only add, never overwrite.
3. **跳过 Visual QA** — First render ALWAYS has overflow/alignment issues. Run at least one fix-verify cycle.
4. **凭空编造数字** — IC Memo numbers will be challenged. Every figure must trace to a source file.
5. **中英混排 `.font.name = "Calibri"`** — Must inject `a:ea` + `a:cs` XML for CJK.
6. **页内太多色** — Red + Gold + two greys is enough. Adding blue/green breaks the brand.
7. **决策页自己下结论** — IC Memo presents objectively. Don't vote for the investment committee.
