# Cathay PPT Template v2 — 设计文档

> 日期: 2026-04-28
> 状态: 设计阶段
> 基于: cathay-ppt-template skill v1 (~4,500 lines lib code)

---

## 1. 动机

当前 v1 skill 功能完整但存在三个结构性问题：

1. **文件架构混乱** — `text_engine.py` 1255 行承载 12 个不相关功能域（字体/文本/表格/图表/验证/合并/安全布局/反腐蚀）；`slide_templates.py` 1264 行 16 个模板含大量重复 boilerplate；SKILL.md 和 lib/ 之间有冗余代码
2. **排版质量平淡** — 模板偏表格/图表填充，缺少现代 deck 的留白节奏、层级呼吸感、卡牌式信息设计
3. **维护成本高** — 常量三处定义且数值不一致（SKILL.md CL=10mm vs constants.py CL=11mm），模板参数顺序不统一

**目标**: 一次彻底升级，让调用方（opencode / claude code）的行为得到更清晰、更美观、更一致的 deck。

---

## 2. 文件架构重构

### 2.1 当前 (v1)

```
lib/
├── constants.py          (234 lines)
├── text_engine.py        (1255 lines) GOD MODULE: fonts + text + tables + charts + validation + merge + safe_layout + anti_corruption
├── slide_templates.py    (1264 lines) 16 templates, heavy duplication
├── qc_automation.py      (1052 lines)
├── data_driven.py        (366 lines)
└── svg_embed.py          (275 lines)
```

### 2.2 目标 (v2)

```
lib/
├── __init__.py            # 便捷 re-export: from cathay_ppt import cards, elements, slides, templates
├── constants.py           # 不变 — 品牌色/网格/CJK 表/guard rail floor 的单一真实来源
├── fonts.py               # ← 从 text_engine 拆出: set_run_font, add_mixed_text, get_char_width, CJK_WIDTH tables (从 constants 移入)
├── text_layout.py         # ← 从 text_engine 拆出: setup_text_frame, format_paragraph, set_square_bullet, add_bullet_content, add_multi_text, calc_text_height, calc_textframe_height
├── elements.py            # ← 从 text_engine 拆出 + 新增 Building Blocks
├── tables.py              # ← 从 text_engine 拆出: add_table, smart_table
├── charts.py              # ← 从 text_engine 拆出: setup_chart_style, safe_chart_insert, insert_chart_image, cathay_bar/line/waterfall/pie_chart
├── slides.py              # ← 从 text_engine 拆出: create_cover_slide, create_content_slide, set_title, set_title_with_conclusion, set_slide_title, add_subtitle, add_source_footer, add_page_number
├── safe_layout.py         # ← 从 text_engine 拆出: safe_textbox, safe_chart_insert
├── validation.py          # ← 从 text_engine 拆出: validate_and_fix, save_with_validation, validate_no_overlap, validate_text_fit, qc_presentation, export_to_pdf
├── merge.py               # ← 从 text_engine 拆出: merge_slides, reorder_slides, clear_slide, full_cleanup, _clean_shape
├── slide_templates.py     # 16 个模板使用 Building Blocks — 每个模板 25-35 行 (从 60-100 行减少)
├── qc_automation.py       # 不变
├── data_driven.py         # 不变
└── svg_embed.py           # 不变
```

### 2.3 拆分原则

- 每个模块 `__all__` 明确导出，消费者用 `from lib.xxx import a, b, c`
- `__init__.py` 提供 `from cathay_ppt import *` 作为便捷入口（按需添加）
- 所有模块保持 `from constants import *` 作为默认，不各自定义常量
- CJK/LATIN 字宽表从 `constants.py` 移到 `fonts.py`（它们是字体引擎的实现细节，不是全局常量）

### 2.4 常量一致性修复

| 常量 | v1 (错误) | v2 (修正) |
|------|-----------|----------|
| CL (content left) | constants.py: **11mm** / SKILL.md: **10mm** | 统一为 **10mm** |
| CT (content top) | constants.py: **31mm** / SKILL.md: **29.2mm** | 统一为 **31mm**（实际生成已经用 31mm，SKILL.md 文档错误） |
| CONTENT_BOTTOM_MM | text_engine: **175mm** / qc_automation: **181mm** | 统一为 **175mm**（content zone bottom） / FOOTER_ZONE_TOP=**181mm**（footer zone 起点） |

### 2.5 SKILL.md 精简

**删除**: 所有内联 Python 函数定义（约 800 行代码块），改为 `见 lib/fonts.py::set_run_font()`

**保留**:
- 品牌标识表
- Quick Start（`from lib import ...` 示例）
- 12 Layout 参考表
- PE Layout 选择指南
- IC Memo 结构表
- 使用示例（调用 lib 函数，不内联函数定义）

**新增**:
- Module Reference 表（每个 lib/ 模块的用途和关键导出）
- Building Blocks API 参考
- 16:9 模式切换指南

---

## 3. 视觉排版升级

### 3.1 Building Blocks

用 5 个 primitive 淘汰模板中的重复代码。每个 block 返回自身 bottom_y_mm 以支持链式布局。

```python
# elements.py 新增

def HeaderBar(slide, x, y, w, h, title, color):
    """彩色标题条 (彩色矩形 + 白色文字覆盖). 返回 bottom_y_mm"""

def ContentPanel(slide, x, y, w, h, items, bg=None):
    """带可选背景色的内容面板. 内部用 smart_textbox."""

def KpiStrip(slide, x, y, kpis):
    """指标行 (等同于 add_kpi_row, 重命名统一风格). 返回 bottom_y_mm"""

def Card(slide, x, y, w, h, header, body, color, text_color=WHITE):
    """完整卡片: HeaderBar + 内容 panel. body 可以是 items list 或 table data 或 chart path"""

def MetricRow(slide, x, y, metrics):
    """指标卡行: 水平排列的 KPI 小卡片(3-5个). 每个含 value + label + 可选趋势箭头"""
```

### 3.2 间距三档体系

```python
# constants.py 新增
GAP_XS = 2   # header 内边距 / icon 间
GAP_SM = 4   # 同组元素间
GAP_MD = 6   # 跨组 / 跨区段
GAP_LG = 10  # 主要板块间呼吸间距
```

现有 `GAP_H=5, GAP_V=3` 改为基于此体系:
```python
GAP_H = GAP_MD   # 列间距 6mm
GAP_V = GAP_SM   # 行间距 4mm
```

### 3.3 卡牌式信息设计

将现金的"bullets + table"两层结构升级为三层:

```
[HeaderBar: 章节标题]
  ├── [Card: 左侧内容]
  │     ├── KPI value (16pt bold maroon)
  │     └── Bullet points (10pt)
  ├── [Card: 右侧图表]
  │     └── Chart image
  └── [Card: 底部结论]
        └── Key takeaway (gold text on light bg)
```

每个 Card 用浅灰背景 `#FAFAFA` + 1pt maroon 左边框。视觉上有呼吸感但信息密度不减。

### 3.4 16:9 Grid 支持

```python
# grid.py (新增)
class SlideGrid:
    """支撑 4:3 和 16:9 两种格式的 grid 系统"""
    
    def __init__(self, aspect="4:3"):
        if aspect == "16:9":
            self.slide_w, self.slide_h = 254, 143  # mm
            self.ct, self.cb, self.cl, self.cw = 31, 131, 10, 233
        else:
            self.slide_w, self.slide_h = 254, 190.5
            self.ct, self.cb, self.cl, self.cw = 31, 181, 10, 233
        self.ch = self.cb - self.ct
        # 预计算所有 grid 常量...
    
    def col(self, n, of_total):  # 返回 (x_mm, w_mm) 如 grid.col(1, 3) → 第一列
    def row(self, n, of_total):  # 返回 (y_mm, h_mm)
    
    def gap_h(self, size_key="md"):  # 按三档间距返回
    def gap_v(self, size_key="md"):
```

模板函数用 `slide_grid = SlideGrid("4:3")` 而非手写的 X1/X2_HALF 等。切到 16:9 只需改构造函数参数。

### 3.5 表格质量细节

- 行线：`0.3pt #D9D9D9`（当前无行线，视觉上很散）
- 表头底线：`1pt #800000` maroon
- 无竖线（已有）
- 交替行背景：保留 `#F2F2F2`
- 数字列右对齐强制设置 `p.alignment = PP_ALIGN.RIGHT`（当前所有列都是左对齐）

### 3.6 字号层级重新校准

| 层级 | v1 | v2 | 变化 |
|------|----|----|------|
| 标题 | 20-28pt | 20-24pt Bold | 收窄上限 |
| 副标题 | 14pt | 12pt Semibold | 更克制 |
| 段落主题 (level 0) | size+2 (12pt) | 12pt Bold Maroon | 一致 |
| 正文 (level 1) | 10pt | 10pt | 不变 |
| 子级 (level 2) | 8-9pt | 9pt | 地板提至 9pt |
| Source footer | 7pt | 7pt | 不变 |

---

## 4. 模板架构升级

### 4.1 重构策略

所有 16 个模板用 Building Blocks 重写。目标: 每个模板 25-35 行 (从 60-100 行减少)。

**前 (v1)** — `template_sidebar_case_study` 约 42 行，含手动 sidebar 布局逻辑

**后 (v2)** — 约 28 行:
```python
def template_sidebar_case_study(prs, title, subtitle, sidebar_metrics,
                                main_items, bottom_table, source):
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)
    content_bottom = CB if not bottom_table else CB - 45

    Card(slide, X1, CT, ONE_QUARTER, content_bottom - CT,
         header=None, body=sidebar_metrics,
         color=CATHAY_RED, text_color=CATHAY_WHITE, sidebar_mode=True)

    ContentPanel(slide, X2_Q34, CT, THREE_QUARTER, content_bottom - CT, main_items)

    if bottom_table:
        smart_table(slide, bottom_table, top_mm=content_bottom + GAP_SM)

    add_source_footer(slide, source)
    return slide
```

### 4.2 参数标准化

所有 16 个模板签名规范化为:

```python
def template_xxx(prs, title, subtitle, *, source="", **data_kwargs):
```

- `source` 是 keyword-only 参数（允许默认值 "" 在独立使用时跳过）
- `**data_kwargs` 携带模板特定数据（如 `kpis`, `bullets`, `table_data` 等）
- 调用者不需要知道哪些参数是哪个模板的 — `render_spec()` 从 spec dict 解包并传递

### 4.3 Universal Template & 路由简化

```python
# data_driven.py 中

def template_universal(prs, spec):
    """单一入口: 接受 spec dict, 自动路由到正确的子模板"""
    t = spec['template']
    renderer = TEMPLATE_ROUTER[t]
    title = spec.get('title', '')
    subtitle = spec.get('subtitle', '')
    source = spec.get('source', '')
    data = spec.get('data', {})
    return renderer(prs, title, subtitle, source=source, **data)
```

`render_spec()` 简化为调用 `template_universal()`，消除当前 16 路 if/else 或 dict lookup 的手动参数打包逻辑。

### 4.4 新增 3 个模板

| # | 名称 | 用途 | 优先级 |
|---|------|------|--------|
| T17 | `template_timeline` | 里程碑时间线（横轴时间 + 纵轴事件卡） | P1 |
| T18 | `template_cap_table` | 股权结构表（投资人/轮次/持股/估值 四列） | P2 |
| T19 | `template_number_story` | 大数字叙事（3-4 个超大 KPI + 简短解释） | P1 |

---

## 5. 实施计划

### Phase 1: 模块拆分 (预计 1 次实施，关键路径)
1. 从 `text_engine.py` 拆分出 `fonts.py`, `text_layout.py`, `elements.py`, `tables.py`, `charts.py`, `slides.py`, `safe_layout.py`, `validation.py`, `merge.py`
2. 修复常量一致性 (constants.py)
3. 更新所有 import 路径
4. 验证: 运行 `python -c "from lib.fonts import *; from lib.elements import *; ..."` 无导入错误
5. 确保现有 `slide_templates.py` 和 `qc_automation.py` 的 import 都正确迁移

### Phase 2: Building Blocks + 模板重构 (预计 1 次实施)
1. 在 `elements.py` 中实现 5 个 Building Blocks，每个含完整 docstring 和 bottom_y_mm 返回
2. 用 Building Blocks 重写 16 个模板（每模板 25-35 行）
3. 新增 T17 (timeline), T19 (number_story)
4. 实现 `SlideGrid` 类，更新 grid 常量
5. 实现三档间距体系
6. 验证: 用 `data_driven.py` 构建一个示例 deck，视检 PDF

### Phase 3: SKILL.md 精简 + 表格质量 (预计 1 次实施)
1. 删除 SKILL.md 中所有内联代码
2. 新增 Module Reference 表、Building Blocks API 参考
3. 表格行线质量提升
4. 字号层级校准

### Phase 4: 验证 & 灰度
1. 用现有 equity-research skill 调用 v2 生成一个完整 deck
2. 视检 PDF 输出
3. QC pipeline 全量通过
4. T18 (cap_table) 实现

---

## 6. 风险评估

| 风险 | 概率 | 影响 | 缓解 |
|------|------|------|------|
| backward compat 破坏 | 中 | 高 | 保留旧函数名作为 alias, `__init__.py` 提供兼容导入 |
| 16:9 grid 计算精度 | 低 | 中 | 用已知 4:3 值作为基准验证 |
| Building Blocks 不够灵活 | 低 | 低 | 每个 block 暴露 `**kwargs` 透传到底层 python-pptx API |
| 模板重写引入视觉回归 | 低 | 中 | Phase 4 必须生成 side-by-side 1-page comparison |

---

## 7. 待定 (Open Questions)

- [ ] T18 (cap_table) 是否需要？还是用现有的 smart_table 已够？
- [ ] `__init__.py` 应该暴露多宽？(`from cathay_ppt import *` 全量暴露 vs `from cathay_ppt.elements import Card`)
- [ ] 是否需要保留 `.inch` 版常量（CONTENT_LEFT, CONTENT_TOP 等）还是全部 migrate 到 mm？
- [ ] Chart 生成函数（cathay_bar_chart 等）应该放在 `charts.py` 还是独立成 `chart_templates.py`？
- [ ] 实施时需同步更新 `references/ppt-generation-rules.md`（section 4 的 grid 常量、section 6.6 的 content zone 边界表），消除 CT=29.2mm vs 31mm 的矛盾
- [ ] 实施时需同步更新 `references/text-fitting-engine.md`（字宽/行高表、smart_textbox 调用方式）
