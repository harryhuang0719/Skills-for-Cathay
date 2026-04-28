# PPT Generation Rules

所有PPT生成必须遵守以下规则。违反任何一条都会导致低质量输出。

---

## 1. Chart Insertion — 只指定宽度，不指定高度

```python
# CORRECT — 只指定宽度，高度按原始比例自动计算
slide.shapes.add_picture(chart_path, Mm(x), Mm(y), Mm(width))

# WRONG — 同时指定宽高会拉伸变形
slide.shapes.add_picture(chart_path, Mm(x), Mm(y), Mm(width), Mm(height))
```

**Chart尺寸参考** (matplotlib figsize → PPT插入宽度):
- figsize=(8, 4.5) → 插入宽度 Mm(200)，全宽图表
- figsize=(6, 4.5) → 插入宽度 Mm(150)，半宽图表（用于双栏布局）
- figsize=(7, 5) → 插入宽度 Mm(170)，heatmap等正方形图表

**位置参考**:
- 全宽图表: x=Mm(15), y=Mm(32)
- 左半图表: x=Mm(10), y=Mm(32)
- 右半图表: x=Mm(130), y=Mm(32)

---

## 2. 字体处理 — 必须用 set_run_font() + 拆分中英文 runs

**铁律**: cathay-ppt-template的`set_run_font()`是唯一正确的字体设置方式。它通过XML设置`a:ea` typeface和`a:altLang`属性。

### macOS 字体安装要求

**KaiTi（楷体）在macOS上不是系统自带字体**，需要手动安装：
- macOS自带的是 **STKaiti**（华文楷体）和 **Kaiti SC**（楷体-简），这些是不同的字体
- Windows的 **KaiTi**（楷体）来自中易电子，内部字体名为 `KaiTi` / `楷体`
- 必须安装 `simkai.ttf` 到 `~/Library/Fonts/` 才能正确渲染
- 安装后通过 `fc-match "KaiTi"` 验证是否解析正确
- CJK检测范围必须包含全角字符: `'\u4e00' <= c <= '\u9fff' or '\u3000' <= c <= '\u303f' or '\uff00' <= c <= '\uffef'`

```python
# 引入cathay-ppt-template的完整helper函数集
# （不要自己重写简化版）

from pptx.oxml.ns import qn
from lxml import etree

def set_run_font(run, text, size_pt=None, bold=False, color_rgb=None):
    size_pt = size_pt or 10.5
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color_rgb:
        run.font.color.rgb = color_rgb
    # 扩展CJK检测范围：包含CJK汉字、标点、全角字符
    has_chinese = any('\u4e00' <= c <= '\u9fff' or '\u3000' <= c <= '\u303f' or '\uff00' <= c <= '\uffef' for c in text)
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

**混合中英文文本 — 必须拆分为多个runs**:
```python
# CORRECT — 拆分runs
import re

def add_mixed_text(para, text, size_pt=10.5, bold=False, color_rgb=None):
    """将混合中英文文本拆分为多个runs，每个run单独设置字体。"""
    # 按照中文/非中文边界拆分
    segments = re.findall(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]+|[^\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]+', text)
    for seg in segments:
        if seg.strip() or seg == ' ':
            run = para.add_run()
            set_run_font(run, seg, size_pt=size_pt, bold=bold, color_rgb=color_rgb)

# WRONG — 单个run包含中英文混合
run.text = "季度收入 $13.64B"  # KaiTi或Calibri只能选一个
```

### Font Setting Iron Rule

**绝对禁止直接设置 `run.font.name` 或 `para.font.name`**。所有文本设置必须通过:
- `set_run_font(run, text, ...)` — 单语言文本
- `add_mixed_text(para, text, ...)` — 混合中英文文本

这包括: table cells, sidebar text, chart annotations, KPI labels, section divider titles.

**违反检查**: 在PPT生成脚本中搜索 `.font.name =`，除了 `set_run_font` 函数内部，不应出现其他地方。

---

## 3. 内容深度 — 每页至少200字

**铁律**: 不允许"3个bullet point"式的浅薄slide。每页必须有实质性分析。

**内容结构模板**:
```
[段落1: 核心观点 + 数据支撑] (50-80字)
[段落2: 展开论证 + 因果分析] (60-100字)
[段落3: 对比/转折 + 启示] (50-80字)
[段落4: 结论/展望] (40-60字)
```

**各类slide的最低字数**:
| Slide类型 | 最低字数 | 内容要求 |
|-----------|---------|---------|
| Investment Summary | 250 | KPI + 完整thesis段落 |
| Company Overview | 300 | 业务描述 + 竞争定位 + 关键驱动 |
| Industry/Competition | 250 | TAM + 结构 + 动态 + 定位 |
| Financial Analysis | 150 + 数据表 | 趋势解读 + 异常标注 |
| Valuation | 200 + 数据表 | 方法论 + 假设 + 结果 + 敏感性 |
| Bull/Bear Case | 300每侧 | 论点 + 催化剂 + 反驳 |
| CIO Verdict | 350 | 独立判断 + 评估 + 决策 + 行动计划 |
| Risk Analysis | 200 + 风险表 | 风险分类 + 概率/影响 + 缓解 |

---

## 4. 布局多样性 — 至少5种不同Grid Pattern

**铁律**: 不允许连续3页以上使用相同布局。25页deck至少用5种不同grid pattern。

**推荐的Slide类型 → Grid映射**:

| Slide类型 | Grid Pattern | cathay-ppt-template常量 |
|-----------|-------------|----------------------|
| Cover | Layout [0] — 全屏 | — |
| Investment Summary | KPI row + full body | `add_kpi_row()` + body below |
| Company Overview | 1/4 dark sidebar + 3/4 | `X1, ONE_QUARTER` + `X2_Q34, THREE_QUARTER` |
| Revenue/Segments | 1/2 + 1/2 | `X1, HALF` + `X2_HALF, HALF` |
| Industry Analysis | 2/3 text + 1/3 chart | `X1, TWO_THIRDS` + `X2_T23, ONE_THIRD` |
| Financial Tables | Full width | `X1, CW` |
| Margin/EPS Charts | Chart top + text bottom | chart at CT + text at ~Mm(110) |
| Valuation Comps | 1/3 metrics + 2/3 table | `X1, ONE_THIRD` + `X2_T23, TWO_THIRDS` |
| Bull vs Bear | 1/2 + 1/2 | `X1, HALF` + `X2_HALF, HALF` |
| CIO Verdict | 1/4 sidebar + 3/4 | 左侧深色 verdict panel + 右侧 reasoning |
| Scenarios | 1/3 + 1/3 + 1/3 | `X1, THIRD` + `X2_MID, THIRD` + `X3_RIGHT, THIRD` |
| Risk Matrix | Full width table | `add_table()` |
| Section Divider | Layout [11] | white on dark |

**Grid常量 (从cathay-ppt-template v2)**:
```python
CL=10; CT=31; CW=234; CH=150; GAP_H=6; GAP_V=4
HALF=(CW-GAP_H)/2          # 114.0mm
THIRD=(CW-GAP_H*2)/3       # 74.0mm
ONE_THIRD=(CW-GAP_H)*1/3   # 76.0mm
TWO_THIRDS=(CW-GAP_H)*2/3  # 152.0mm
ONE_QUARTER=(CW-GAP_H)*1/4 # 57.0mm
THREE_QUARTER=(CW-GAP_H)*3/4 # 171.0mm
X1=CL                       # 10mm
X2_HALF=CL+HALF+GAP_H       # 130.0mm
X2_T23=CL+ONE_THIRD+GAP_H   # 92.0mm
X2_Q34=CL+ONE_QUARTER+GAP_H # 73.0mm
```

---

## 5. 必须使用的 cathay-ppt-template Helper Functions

**不允许自己从头写简化版。必须完整复制以下函数到PPT生成脚本中：**

| Function | 用途 | 必须用 |
|----------|------|--------|
| `setup_text_frame(tf)` | 设置margins, autosize, word_wrap | 每个TextBox |
| `format_paragraph(para)` | 设置spacing, indent, line_spacing | 每个段落 |
| `set_run_font(run, text, ...)` | 字体+中英文检测+XML属性 | 每个run |
| `set_square_bullet(para)` | PPT原生方块bullet | 有bullet的段落 |
| `add_bullet_content(tf, items)` | 批量添加分层内容 | 结构化内容 |
| `add_source_footer(slide, text)` | 8pt source在y=18cm | 每页 |
| `add_table(slide, ...)` | Cathay格式化表格 | 数据表 |
| `add_kpi_row(slide, kpis)` | KPI指标行 | Summary页 |
| `safe_chart_insert(slide, ...)` | 图表插入+返回bottom_y | 每个图表 |
| `safe_textbox(slide, ...)` | 创建自适应高度文本框 | 每个内容文本框 |
| `validate_text_fit(pptx_path)` | 检查文本是否溢出 | QC阶段 |
| `set_title_with_conclusion(slide, topic, conclusion)` | 结论式标题 | Layout [4]每页 |
| `add_section_marker(slide, x, y, icon)` | 段落视觉标记 | 有2+段落主题的页面 |

---

## 6. Text Overflow Prevention — 最重要的规则

**MUST READ**: `references/text-fitting-engine.md` — 精确文字高度计算、auto-fit textbox、smart table。

**铁律**: python-pptx没有文字渲染引擎。每创建一个textbox或table，都必须验证text是否fit。

### 关键常量
- CJK字符宽度(10pt楷体): **3.15mm/字**
- Latin字符宽度(10pt Calibri): **2.0mm/字**
- 10pt + 1.2x行间距 = 每行高 **4.23mm**
- Source footer: **7pt, 5mm height, y=182mm** (永远不变)
- Content zone: **31mm to 181mm** (150mm可用)
- 表格最小row height: **7mm** (9pt CJK) 或 **6mm** (8pt CJK)

### 创建textbox前必须检查
```python
# 永远不要:
slide.shapes.add_textbox(Mm(x), Mm(y), Mm(w), Mm(h))  # 盲猜h
# 永远要:
smart_textbox(slide, x, y, w, items, max_bottom_mm=181)  # auto-fit
```

### 创建table前必须检查
```python
# 永远不要:
add_table(slide, data, row_height=6)  # 6mm装不下9pt中文
# 永远要:
smart_table(slide, data, font_size=9, min_row_h=7)  # auto-fit
```

### 保存前必须validate
```python
# 每个脚本结尾:
fixes = validate_and_fix(prs)
prs.save(path)
```

---

## 6.5 Slide Merge — 必须用内置merge函数

**铁律**: 合并多个单slide文件时，必须用`merge_slides()`处理image rId映射。手写copy会丢图片。

```python
from text_fitting_engine import merge_slides
merge_slides(
    slide_files={1: 'slide_01.pptx', 2: 'slide_02.pptx', ...},
    output_path='final.pptx',
    slide_order=[1, 2, 3, ...]
)
```

---

## 6.6 Content Zone Boundaries — 硬性边界

| 区域 | 位置 | 说明 |
|------|------|------|
| Title bar | 0 - 20mm | 模板自带，不要放content |
| Subtitle zone | 21 - 30mm | 金色副标题（可选）|
| **Content zone** | **31mm - 181mm** | 所有内容必须在此范围 |
| Source footer | 182mm | 7pt, 5mm height |
| Page number | 182mm | 7pt, right-aligned |
| Slide bottom | 190.5mm | 不要超过 |

**检查方法**: 每个shape的 `top + height` 不得超过 181mm（source/page除外）。

---

## 7. QC Checklist (生成后必查)

- [ ] 打开PPT检查所有图表是否变形（宽高比是否正确）
- [ ] 抽查3页: 中文是否为楷体(KaiTi)，英文/数字是否为Calibri（macOS需确认simkai.ttf已安装）
- [ ] 统计内容页字数: 是否每页≥200字
- [ ] 统计布局类型: 是否≥5种不同grid pattern
- [ ] 运行 `validate_no_overlap()` 检查shape重叠
- [ ] PDF导出后检查渲染是否正确
- [ ] 确认所有slide有source footer
- [ ] 运行 `validate_text_fit()` 无 warnings
- [ ] 标题样式一致: 模式A用`set_title_with_conclusion()`或模式B用`add_slide_title()`（标题包含核心结论+小标题）
- [ ] 内容 slides 使用 level-0/1/2 bullet 层次（3+ topics 必须有 level-0 headers）
- [ ] 有 2+ level-0 headers 的页面有 section marker icons
- [ ] **内容撑满**: 每页content zone无大面积空白，表格+bullets+视觉元素铺满到source footer区域
- [ ] **字体颜色**: 深色背景=白色文字，浅色背景=深红标题+黑色正文（绝不在浅色背景用金黄色标题）
- [ ] **CJK检测**: set_run_font中的CJK范围包含全角字符(\u3000-\u303f, \uff00-\uffef)

---

## 7. Overflow Prevention — safe_chart_insert() + safe_textbox()

**铁律**: 使用 `safe_chart_insert()` 替代 `slide.shapes.add_picture()` 插入图表。使用 `safe_textbox()` 替代 `slide.shapes.add_textbox()` 创建文本框。

**Chart + Text stacking pattern** (最常用):
```python
# Chart on top, text below — dynamic height calculation
chart_bottom = safe_chart_insert(slide, chart_path, x_mm=CL, y_mm=CT, w_mm=220)
text_top = chart_bottom + GAP_V
_, tf = safe_textbox(slide, CL, text_top, CW)  # auto-fills to CONTENT_BOTTOM_MM
add_bullet_content(tf, items)
add_source_footer(slide, "Source: ...")
```

**Content Zone Constants**:
- Title bar ends at: y = 29.2mm (CT)
- Source footer starts at: y = 180mm
- Usable content zone: y = 29.2mm to y = 175mm (CONTENT_BOTTOM_MM)
- Safety margin: keep 5mm above footer

**Never hardcode text box heights for variable content**. Always use `safe_textbox()` with no `h_mm` to auto-fill remaining space.

**Side-by-side charts**:
```python
# Two charts side by side, text below both
left_bottom = safe_chart_insert(slide, chart_left, x_mm=CL, y_mm=CT, w_mm=HALF)
right_bottom = safe_chart_insert(slide, chart_right, x_mm=X2_HALF, y_mm=CT, w_mm=HALF)
text_top = max(left_bottom, right_bottom) + GAP_V
_, tf = safe_textbox(slide, CL, text_top, CW)
```

---

## 8. Content Hierarchy — 标准化Bullet结构

**铁律**: 所有内容slide使用 `add_bullet_content()` 的3级层次。不允许使用plain paragraph列表。

**标准层次结构**:

| Level | 格式 | 用途 | 示例 |
|-------|------|------|------|
| 0 | Bold red (#800000), size+2pt, **no bullet**, no indent | 段落主题/子标题 | "收入加速驱动力" |
| 1 | Filled ■ bullet, standard size, 0.5cm indent | 关键论点 | "HBM收入QoQ+40%，ASP上行" |
| 2 | Filled ■ bullet, size-1pt, grey (#808080), 1.0cm indent | 补充数据/细节 | "FQ1 HBM $1.5B vs FQ4 $1.1B" |

**何时使用哪种结构**:
- **bullet-hierarchy**: 3+ distinct topics → 必须用level-0 headers分隔（大多数分析slides）
- **prose**: Thesis/CIO/Bull/Bear等需要完整论证的段落（用`add_mixed_text`）
- **mixed**: 表格/图表旁的解读（level-0 header + prose paragraph）

**items 构造模板**:
```python
items = [
    ("收入加速驱动力", 0),                    # section header
    ("HBM收入QoQ+40%，ASP持续上行", 1),         # key point
    ("FQ1 HBM收入$1.5B vs FQ4 $1.1B", 2),     # data support
    ("DDR5升级周期驱动bit shipment增长", 1),
    ("利润率结构性扩张", 0),                    # new section
    ("GM从32%扩张至56%非周期性因素", 1),
    ("HBM mix带来的永久性利润率提升", 2),
]
add_bullet_content(tf, items)
```

**Section Icons**: 有2+个level-0 headers的slide，在每个level-0左侧放置4mm section marker:
```python
icons = auto_assign_icons(items)  # keyword-based assignment
# Or manual: icons = {"收入加速驱动力": ICON_FINANCIAL, "利润率结构性扩张": ICON_FINANCIAL}
```

---

## 9. Slide Titles — 两种标题模式

### 模式A: 红色标题栏模式（默认，适用于IC Memo等正式文件）

使用模板Layout [4]的红色标题栏 + `set_title_with_conclusion()`。

**格式**: `[Topic] — [Key Conclusion]`
- Topic: 白色, 20pt bold
- Separator: " — " 金色
- Conclusion: 金色, 18pt bold

```python
set_title_with_conclusion(slide, "投资摘要", "BUY, 目标价$520 (+14%)")
```

### 模式B: 深红文字标题模式（推荐，适用于项目介绍/pitch deck）

使用空白Layout [9]，自定义标题区域。标题为**深红色(#800000)大字左对齐**，下方有**灰色小标题**说明核心内容，再加一条**红色分隔线**。

**关键规则**:
- 标题: 深红色(#800000), 22pt, bold, 左对齐
- 小标题: 深灰色(#404040), 11pt, 左对齐，解释该页核心结论
- 分隔线: 红色(#800000), 0.8mm高, 全宽
- 标题区域占用 5mm-24mm，内容从 26mm 开始

```python
def add_slide_title(slide, title, subtitle):
    """深红色文字标题 + 灰色小标题 + 红色分隔线"""
    # Title: dark red, left-aligned
    tx = slide.shapes.add_textbox(Mm(CL), Mm(5), Mm(CW), Mm(10))
    tf = tx.text_frame; setup_text_frame(tf)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    add_mixed_text(p, title, sz=22, bold=True, color=CATHAY_RED)  # #800000
    # Subtitle: dark grey
    tx2 = slide.shapes.add_textbox(Mm(CL), Mm(16), Mm(CW), Mm(7))
    tf2 = tx2.text_frame; setup_text_frame(tf2)
    p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.LEFT
    add_mixed_text(p2, subtitle, sz=11, color=RGBColor(0x40,0x40,0x40))
    # Red divider line
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(CL), Mm(24), Mm(CW), Mm(0.8))
    ln.fill.solid(); ln.fill.fore_color.rgb = CATHAY_RED; ln.line.fill.background()

# 使用空白layout
slide = prs.slides.add_slide(prs.slide_layouts[9])  # Vide — completely blank
add_slide_title(slide, "投资摘要", "三重催化共振期——深海智人是国内最具差异化的深海ROV标的")
# 内容从 y=26mm 开始
```

**何时选择哪种模式**:
- **模式A (红色标题栏)**: IC Memo、FDD Summary、正式投委材料
- **模式B (深红文字标题)**: 项目介绍、pitch deck、行业研究、对外展示材料

**共同规则**: 标题不是通用描述，而是该页的核心结论。

```python
# CORRECT
add_slide_title(slide, "深海行业与竞争格局", "全球寡头垒60年护城河，中国创业公司四大窗口")
# WRONG
add_slide_title(slide, "行业概述", "")  # 通用描述，没有结论
```

---

## 10. 内容撑满 — 页面不留大面积空白

**铁律**: 每页内容必须从标题区域下方一直铺满到source footer区域(182mm)。不允许内容只占半页。

**撑满的定义**:
- content zone (26mm-181mm，模式B；或31mm-181mm，模式A) 应被表格、文本框、图形元素充分利用
- 如果表格+bullets只占半页，需要：①增加更多分析数据 ②添加callout box/sidebar等视觉元素 ③拆分为左右两栏填满
- 视觉元素可用于丰富布局：dark red sidebar、gold accent bars、callout boxes(红色左边框+浅粉背景)、horizontal divider lines

**丰富布局的视觉元素**:

```python
# Dark red sidebar (白色文字)
def add_sidebar(slide, x, y, w, h, items):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(x), Mm(y), Mm(w), Mm(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = CATHAY_RED
    # overlay textbox with white text inside

# Gold vertical accent bar
def add_vbar(slide, x, y, h, w=3):
    b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(x), Mm(y), Mm(w), Mm(h))
    b.fill.solid(); b.fill.fore_color.rgb = CATHAY_GOLD

# Callout box (浅粉背景 + 红色左边框)
def add_callout(slide, x, y, w, h, text):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(x), Mm(y), Mm(w), Mm(h))
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(0xF8,0xF0,0xF0)
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(x), Mm(y), Mm(2.5), Mm(h))
    accent.fill.solid(); accent.fill.fore_color.rgb = CATHAY_RED

# Horizontal divider line
def add_hline(slide, x, y, w):
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(x), Mm(y), Mm(w), Mm(0.5))
    ln.fill.solid(); ln.fill.fore_color.rgb = CATHAY_LTGREY
```

---

## 11. 字体颜色规则 — 深色背景白字，浅色背景深红标题

**铁律**: 字体颜色取决于背景色。

| 背景 | 标题/Header | 正文 | 子级文字 |
|------|-----------|------|---------|
| 白色/浅色背景 | 深红色 #800000, bold | 黑色 #000000 | 灰色 #808080 |
| 深红色背景(sidebar/KPI) | 白色 #FFFFFF, bold | 白色 #FFFFFF | 浅金色 #E8D590 |
| 表格header行(#800000) | 白色 #FFFFFF, bold | — | — |

**绝对禁止**: 深色背景上使用黑色/深色文字, 浅色背景上使用金黄色/居中标题。

---

## 12. Unisun模板使用 — 左侧红竖线Layout

**推荐模板**: `~/Desktop/Cathay PPT Template/Unisun - Investment Memo.pptx`

此模板的Layout[4] (`5_Red Slide`) 内置左侧深红竖线(Rectangle 9, 5mm宽×7.5"高)和Cathay logo。比默认模板更专业。

**关键参数**:
- 模板尺寸: 10.00" x 7.50" (254mm x 190.5mm)
- 红色竖线: x=0, width=5mm, 全高
- Title placeholder: left=10.6mm, top=2.9mm (隐藏后用自定义标题)
- Content left: **12mm** (红竖线5mm + 间距7mm)
- Content width: **236mm** (254 - 12 - 6)
- Layout[0] = Cover, Layout[4] = Content, Layout[11] = Section divider

**使用方法**:
```python
TEMPLATE = "~/Desktop/Cathay PPT Template/Unisun - Investment Memo.pptx"
prs = Presentation(TEMPLATE)
# 删除52张已有slides
while len(prs.slides) > 0:
    rId = prs.slides._sldIdLst[0].rId
    prs.part.drop_rel(rId); del prs.slides._sldIdLst[0]
# 内容slide使用Layout[4]
slide = prs.slides.add_slide(prs.slide_layouts[4])
# 隐藏模板标题placeholder，使用自定义深红标题
```

---

## 13. IB视觉元素库 — 避免overlap的设计模式

从本session的多次迭代中总结的正确实现方式：

### Metric Card（无overlap版）
```python
def add_metric_card(slide, x, y, w, h, val, lbl, accent_color):
    # 单个card背景 + border（不再用accent line覆盖在card上）
    bg = slide.shapes.add_shape(ROUNDED_RECTANGLE, Mm(x), Mm(y), Mm(w), Mm(h))
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(0xF5,0xF5,0xF5)
    bg.line.width = Pt(1.5); bg.line.color.rgb = accent_color
    # 单个textbox包含value+label（不要分两个textbox）
    tx = slide.shapes.add_textbox(Mm(x+1), Mm(y+2), Mm(w-2), Mm(h-4))
    # value paragraph + label paragraph 在同一个textframe中
```

### Numbered Circle（放在左margin，不与内容overlap）
```python
def add_numbered_circle(slide, x, y, number, color, size=5):
    # 放在x=7mm（红竖线5mm之后，内容区12mm之前）
    # size最大5mm，确保7+5=12mm不超过CL
    circle_x = 7; actual_size = min(size, 5)
    sh = slide.shapes.add_shape(OVAL, Mm(circle_x), Mm(y+1), Mm(actual_size), Mm(actual_size))
```

### Banner Box（P0/P1/P2等section header）
```python
# CORRECT — 全宽banner box，不与circle/accent bar叠放
add_banner_box(slide, right_x, p0_y, right_w, 7, "P0: 投资决策前必须完成", C_RED)
# WRONG — circle + banner 并排放（circle会与banner overlap）
add_numbered_circle(slide, right_x+1, p0_y, "P0", C_RED)  # overlaps with banner below
add_banner_box(slide, right_x+13, p0_y, 55, 7, "投资决策前必须完成", C_RED)
```

### Accent Bar位置规则
- **Gold accent bar**: 只放在textbox的正左侧，且textbox的left必须 > bar的left+width+2mm
- **不要把accent bar和circle/numbered icon放在同一列**——必然overlap
- Accent bar宽度: 3mm；与内容间距: ≥2mm

### Callout Box（用ROUNDED_RECTANGLE）
```python
# CORRECT — 圆角矩形callout
bg = slide.shapes.add_shape(ROUNDED_RECTANGLE, ...)
# accent bar在callout内部左边缘(x相同，width=2.5mm)
# textbox在accent bar右侧(x+5mm)
```

---

## 14. Overlap Prevention Checklist — 每次生成后必查

**铁律**: 所有overlap必须在生成脚本中设计时消除，不能依赖"视觉上看不出来"。

**常见overlap来源及修复**:

| 问题 | 原因 | 修复 |
|------|------|------|
| Metric card accent line vs card bg | 两个shape在同一位置 | 改为card border颜色代替accent line |
| Numbered circle vs banner box | circle太大超出margin | 缩小到5mm，放在x=7mm |
| Accent bar vs content text | bar和textbox在同一列 | 确保textbox.left > bar.left + bar.width + 2mm |
| P0/P1/P2 circle vs accent bar | circle放在accent bar内部 | 去掉accent bar，用banner box代替 |
| Template title placeholder vs custom title | 隐藏的placeholder仍占空间 | 这个overlap是预期的，QC时filter掉 |

**QC脚本中过滤模板placeholder overlap**:
```python
issues = validate_no_overlap(pptx_path)
real_issues = [i for i in issues if 'Title 1' not in i]
print(f'Real overlaps: {len(real_issues)}')  # 应该为0
```

---

## 15. Line Spacing与排版 — 紧凑但可读

**推荐参数**:
- Line spacing: **120%** (120000) — 紧凑但不拥挤，适合信息密集的IB deck
- Spacing before: **3pt** (300) — 段落间呼吸
- Spacing after: **0pt**
- 130%用于行间距宽松的pitch deck，120%用于信息密集的IC memo

**字号层级(严格遵守)**:
| 层级 | 字号 | 用途 |
|------|------|------|
| Title | 20-22pt | 页面标题(深红色bold) |
| Subtitle | 10-11pt | 页面小标题(深灰色) |
| Section header | 10-11pt | Level-0 section headers(深红色bold) |
| Body | 9-10pt | 正文bullets |
| Sub-body | 7.5-8pt | Level-2 sub-bullets(灰色) |
| Caption/Source | 7pt | Source footer |
| KPI value | 16-18pt | Metric card数字(深红色bold) |
| KPI label | 7.5-8pt | Metric card标签(深灰色) |
| Banner text | 8-9pt | Banner box内文字(白色bold) |

---

## 16. 设计资源参考

详见 `references/awesome-design-resources.md`，包含：
- 配色工具(Coolors/Color Hunt)
- 字体搭配(Font Pair/Type Scale)
- 图标资源(Noun Project/Flaticon)
- 布局灵感(Dribbble/Behance搜"pitch deck")
- IB deck设计8原则(视觉层级/色彩克制/留白/一致性/数据可视化/shape语言/icon用法/对比度)
