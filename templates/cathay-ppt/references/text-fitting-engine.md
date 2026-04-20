# Text Fitting Engine

python-pptx没有文字渲染引擎——不知道text实际占多高。这是PPT生成中text overflow的根本原因。

本文档定义了精确的文字高度计算方法和auto-fit机制，所有PPT生成脚本必须使用。

---

## 1. CJK字宽常量表

通过实际PPT渲染测量得出（Calibri/楷体，单位：mm/字符）：

```python
# 每个字符的水平宽度（mm），按字号
CJK_CHAR_WIDTH = {
    7:    2.2,   # 7pt楷体
    7.5:  2.4,
    8:    2.5,
    8.5:  2.7,
    9:    2.85,
    9.5:  3.0,
    10:   3.15,  # 10pt楷体
    10.5: 3.3,
    11:   3.5,
    12:   3.8,
}

LATIN_CHAR_WIDTH = {
    7:    1.4,   # 7pt Calibri
    7.5:  1.5,
    8:    1.6,
    8.5:  1.7,
    9:    1.8,
    9.5:  1.9,
    10:   2.0,   # 10pt Calibri
    10.5: 2.1,
    11:   2.2,
    12:   2.4,
}

def get_char_width(font_pt, is_cjk=False):
    """Get character width in mm for given font size."""
    table = CJK_CHAR_WIDTH if is_cjk else LATIN_CHAR_WIDTH
    # Interpolate for non-standard sizes
    pts = sorted(table.keys())
    if font_pt <= pts[0]: return table[pts[0]]
    if font_pt >= pts[-1]: return table[pts[-1]]
    for i in range(len(pts)-1):
        if pts[i] <= font_pt <= pts[i+1]:
            ratio = (font_pt - pts[i]) / (pts[i+1] - pts[i])
            return table[pts[i]] + ratio * (table[pts[i+1]] - table[pts[i]])
    return table[10]  # fallback
```

---

## 2. 精确文字高度计算器

```python
import math, re
from pptx.oxml.ns import qn

def calc_text_height(text_or_paragraphs, box_width_mm, font_pt=10, line_spacing=1.2, margin_mm=4):
    """计算文字渲染高度（mm）。

    Args:
        text_or_paragraphs: str 或 [(text, font_pt, indent_mm), ...] 列表
        box_width_mm: textbox宽度（mm）
        font_pt: 默认字号
        line_spacing: 行间距倍数（1.2 = 120%）
        margin_mm: 上下margins总和（mm）

    Returns:
        float: 预估渲染高度（mm）
    """
    usable_w = box_width_mm - margin_mm  # 水平margins
    if usable_w <= 0: usable_w = 5

    # Normalize input
    if isinstance(text_or_paragraphs, str):
        paragraphs = [(text_or_paragraphs, font_pt, 0)]
    else:
        paragraphs = text_or_paragraphs

    total_h = margin_mm / 2  # top margin

    for i, (text, p_font, indent_mm) in enumerate(paragraphs):
        if not text.strip():
            total_h += 1.5  # empty paragraph spacing
            continue

        effective_w = usable_w - indent_mm
        if effective_w <= 0: effective_w = usable_w

        # Split text into CJK and Latin segments to calculate width accurately
        segments = re.findall(
            r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]+|[^\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]+',
            text
        )

        total_text_width = 0
        for seg in segments:
            is_cjk = any('\u4e00' <= c <= '\u9fff' for c in seg)
            char_w = get_char_width(p_font, is_cjk)
            total_text_width += len(seg) * char_w

        n_lines = max(1, math.ceil(total_text_width / effective_w))
        line_h = p_font * 0.3528 * line_spacing

        # Space before (except first paragraph)
        if i > 0:
            total_h += 1.0  # ~3pt spacing before

        total_h += n_lines * line_h

    total_h += margin_mm / 2  # bottom margin
    return total_h


def calc_textframe_height(text_frame, box_width_mm):
    """从python-pptx TextFrame对象计算渲染高度。"""
    paragraphs = []
    for p in text_frame.paragraphs:
        p_font = 10  # default
        for r in p.runs:
            if r.font.size:
                p_font = r.font.size / 12700
                break

        # Get indent
        indent_mm = 0
        pPr = p._p.get_or_add_pPr()
        marL = pPr.get('marL')
        if marL:
            indent_mm = int(marL) / 36000

        paragraphs.append((p.text, p_font, indent_mm))

    # Get line spacing
    line_sp = 1.2
    if text_frame.paragraphs:
        pPr = text_frame.paragraphs[0]._p.find(qn('a:lnSpc'))
        if pPr is not None:
            spcPct = pPr.find(qn('a:spcPct'))
            if spcPct is not None:
                line_sp = int(spcPct.get('val', '120000')) / 100000

    # Get margins
    m_lr = (text_frame.margin_left or 72000) / 36000 + (text_frame.margin_right or 72000) / 36000
    m_tb = (text_frame.margin_top or 36000) / 36000 + (text_frame.margin_bottom or 36000) / 36000

    return calc_text_height(paragraphs, box_width_mm, line_spacing=line_sp, margin_mm=m_tb)
```

---

## 3. Auto-Fit Textbox

创建textbox时自动计算高度，如果超出则缩字。

```python
def smart_textbox(slide, x_mm, y_mm, w_mm, items, max_bottom_mm=180,
                  start_font=10, min_font=8, line_spacing=1.2):
    """创建auto-fit textbox：先算高度，超出则缩字。

    Args:
        items: [(text, level), ...] 传给 add_bullet_content
        max_bottom_mm: 内容区底部限制
        start_font: 起始字号
        min_font: 最小字号（不会小于此）

    Returns:
        (shape, text_frame, actual_font_pt)
    """
    max_h = max_bottom_mm - y_mm

    # Find the largest font that fits
    chosen_font = start_font
    for try_font in [start_font, start_font - 0.5, start_font - 1,
                     start_font - 1.5, start_font - 2, min_font]:
        if try_font < min_font: try_font = min_font

        # Estimate height at this font
        paras = []
        for text, level in items:
            indent = 5 if level == 1 else (10 if level >= 2 else 0)
            f = try_font + 1 if level == 0 else (try_font - 1 if level >= 2 else try_font)
            paras.append((text, f, indent))

        est_h = calc_text_height(paras, w_mm, line_spacing=line_spacing)

        if est_h <= max_h:
            chosen_font = try_font
            break
        chosen_font = try_font

    # Create textbox with calculated height
    actual_h = min(max_h, est_h + 3)  # +3mm safety
    from pptx.util import Mm
    txBox = slide.shapes.add_textbox(Mm(x_mm), Mm(y_mm), Mm(w_mm), Mm(actual_h))
    tf = txBox.text_frame
    setup_text_frame(tf)
    add_bullet_content(tf, items, size_pt=chosen_font)

    return txBox, tf, chosen_font
```

---

## 4. Smart Table

自动根据内容计算row height。

```python
def smart_table(slide, data, left_mm=None, top_mm=None, width_mm=None,
                max_bottom_mm=180, font_size=9, min_row_h=7):
    """创建auto-fit表格：扫描每cell内容，自动设min row height。

    Returns:
        (table, bottom_y_mm)
    """
    left_mm = left_mm or CL
    top_mm = top_mm or CT
    width_mm = width_mm or CW
    rows = len(data)
    cols = len(data[0]) if data else 0

    # Calculate column widths (equal distribution)
    col_w = width_mm / cols if cols > 0 else width_mm

    # Calculate max lines needed per row
    row_heights = []
    for ri, row_data in enumerate(data):
        max_lines = 1
        for ci, cell_text in enumerate(row_data):
            text = str(cell_text) if cell_text else ""
            if not text: continue
            has_cjk = any('\u4e00' <= c <= '\u9fff' for c in text)
            char_w = get_char_width(font_size, has_cjk)
            usable = col_w - 3  # cell margins
            if usable <= 0: usable = 5
            lines = max(1, math.ceil(len(text) * char_w / usable))
            lines += text.count('\n')
            max_lines = max(max_lines, lines)

        line_h = font_size * 0.3528 * 1.2
        needed_h = max_lines * line_h + 3  # +margins
        row_heights.append(max(needed_h, min_row_h))

    total_h = sum(row_heights)

    # Check if it fits
    if top_mm + total_h > max_bottom_mm:
        # Scale down: reduce font until it fits
        for smaller_font in [font_size - 0.5, font_size - 1, font_size - 1.5, font_size - 2]:
            if smaller_font < 7: break
            font_size = smaller_font
            row_heights_new = []
            for ri, row_data in enumerate(data):
                max_lines = 1
                for ci, cell_text in enumerate(row_data):
                    text = str(cell_text) if cell_text else ""
                    if not text: continue
                    has_cjk = any('\u4e00' <= c <= '\u9fff' for c in text)
                    char_w = get_char_width(smaller_font, has_cjk)
                    usable = col_w - 3
                    if usable <= 0: usable = 5
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

    # Create table using standard add_table (which handles formatting)
    # But use the calculated total height
    return add_table(slide, data, left_mm=left_mm, top_mm=top_mm,
                     width_mm=width_mm, row_height=total_h/rows, font_size=font_size)
```

---

## 5. Pre-Save Validation

保存前自动检查所有shapes。

```python
def validate_and_fix(prs):
    """保存前自动检查并修复所有text overflow和shape overflow。

    Returns:
        list of fix descriptions
    """
    fixes = []

    for slide in prs.slides:
        for shape in slide.shapes:
            top_mm = shape.top / 36000
            height_mm = shape.height / 36000
            width_mm = shape.width / 36000
            bottom_mm = top_mm + height_mm

            if width_mm < 0.5 or height_mm < 0.5: continue

            # 1. Shape exceeds content zone
            is_footer = False
            if shape.has_text_frame:
                txt = shape.text_frame.text.lower()
                if 'source:' in txt or (len(txt) < 10 and '/' in txt):
                    is_footer = True

            if not is_footer and bottom_mm > 181:
                # Cap height
                new_h = 181 - top_mm
                if new_h >= 5:
                    shape.height = int(new_h * 36000)
                    fixes.append(f"CAP: {shape.name} bottom {bottom_mm:.0f}→181mm")

            # 2. Text overflow within textbox
            if shape.has_text_frame and height_mm >= 5:
                est_h = calc_textframe_height(shape.text_frame, width_mm)
                if est_h > height_mm * 1.1:
                    # Reduce font
                    for target in [9, 8.5, 8, 7.5, 7]:
                        for p in shape.text_frame.paragraphs:
                            for r in p.runs:
                                if r.font.size and r.font.size / 12700 > target:
                                    r.font.size = Pt(target)
                        new_est = calc_textframe_height(shape.text_frame, width_mm)
                        if new_est <= height_mm:
                            fixes.append(f"FONT: {shape.name} reduced to {target}pt")
                            break

    return fixes


def save_with_validation(prs, path):
    """保存前自动validation+fix。"""
    fixes = validate_and_fix(prs)
    if fixes:
        print(f"Auto-fixed {len(fixes)} issues before save:")
        for f in fixes[:10]:
            print(f"  {f}")
    prs.save(path)
    return fixes
```

---

## 6. Source Footer标准

**固定规格，不再调整：**

```python
SOURCE_FONT_PT = 7       # 7pt — 永远不变
SOURCE_BOX_HEIGHT_MM = 5  # 5mm — 7pt单行文字刚好fit
SOURCE_Y_MM = 182        # 在content zone(181mm)之外

def add_source_footer(slide, source_text):
    """标准source footer — 7pt, 5mm height, y=182mm."""
    txBox = slide.shapes.add_textbox(Mm(CL), Mm(SOURCE_Y_MM), Mm(180), Mm(SOURCE_BOX_HEIGHT_MM))
    tf = txBox.text_frame
    setup_text_frame(tf)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    set_run_font(run, f"Source: {source_text}", size_pt=SOURCE_FONT_PT, color_rgb=CATHAY_GREY)
```

---

## 7. Slide Merge with Image Support

```python
import copy, io
from pptx.oxml.ns import qn

def merge_slides(slide_files, output_path, template_path=None, slide_order=None):
    """合并多个单slide文件为一个deck，正确处理image relationships。

    Args:
        slide_files: dict {slide_num: path} 或 list of paths
        output_path: 输出文件路径
        template_path: 模板文件路径
        slide_order: list of slide_nums，控制页面顺序
    """
    template_path = template_path or TEMPLATE
    master = Presentation(template_path)

    # Clear template slides
    while len(master.slides) > 0:
        rId = master.slides._sldIdLst[0].rId
        master.part.drop_rel(rId)
        del master.slides._sldIdLst[0]

    if isinstance(slide_files, list):
        slide_files = {i+1: p for i, p in enumerate(slide_files)}

    if slide_order is None:
        slide_order = sorted(slide_files.keys())

    for src_num in slide_order:
        if src_num not in slide_files: continue
        src_path = slide_files[src_num]
        if not os.path.exists(src_path): continue

        src_prs = Presentation(src_path)
        src_slide = src_prs.slides[0]

        # Match layout
        layout_name = src_slide.slide_layout.name
        target_layout = master.slide_layouts[4]  # default
        for layout in master.slide_layouts:
            if layout.name == layout_name:
                target_layout = layout; break

        new_slide = master.slides.add_slide(target_layout)

        # Collect image blobs from source
        img_map = {}
        for rel in src_slide.part.rels.values():
            if "image" in str(rel.reltype):
                try: img_map[rel.rId] = rel.target_part.blob
                except: pass

        # Register images in new slide, build rId mapping
        rId_remap = {}
        for old_rId, blob in img_map.items():
            image_part, new_rId = new_slide.part.get_or_add_image_part(io.BytesIO(blob))
            rId_remap[old_rId] = new_rId

        # Copy shapes with remapped image references
        for shape in src_slide.shapes:
            el = copy.deepcopy(shape._element)
            for blip in el.findall('.//' + qn('a:blip')):
                old_rId = blip.get(qn('r:embed'))
                if old_rId in rId_remap:
                    blip.set(qn('r:embed'), rId_remap[old_rId])
            new_slide.shapes._spTree.append(el)

    master.save(output_path)
    return len(master.slides)
```
