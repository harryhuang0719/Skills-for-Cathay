# Cathay PPT Template v3

Template: `assets/template.pptx` (阿维塔模板, 10.00" × 7.50", 4:3, 12 layouts)

**v3 核心**: 左侧深红竖线(5mm) + dark red 标题 + STKaiti/Calibri 字体栈 + 宽松排版 + Section 驱动的工作流。

---

## Quick Start

```python
import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/cathay-ppt-template/lib"))

from constants import *          # 品牌色/网格/字号/间距
from fonts import set_run_font, add_mixed_text
from text_layout import smart_textbox, add_bullet_content
from tables import smart_table
from slides import create_cover_slide, create_content_slide, create_section_divider, add_source_footer
from elements import Card, KpiStrip, SectionBlock, IconBadge, MetricRow, DividerLine
from slide_templates import *     # 19 个模板函数
```

---

## Brand Identity

| Element | Value |
|---------|-------|
| Primary | `#800000` MAROON |
| Gold accent | `#E8B012` |
| Chinese Font | **STKaiti** (华文楷体) |
| English/Number Font | **Calibri** |
| Title | 20pt Bold dark red, left-aligned |
| Body | **11.5pt** |
| Small/Table | 10pt |
| Source footer | 8pt |
| Line spacing | **1.3×** (130%) |

---

## Layout System

Layout [4] "5_Red Slide" 包含:
- **左侧深红竖线** (5mm 宽, 全高) — 模板自带
- Title placeholder (x=10.7mm, dark red 文字)
- Content zone: **CL=14mm** (红线5mm + 间距9mm), **CT=31mm**, **CB=181mm**, **CW=230mm**

Layout [11] "1_Red Slide" — Section Divider (深红底+白色标题)

---

## Module Reference

| Module | Key Exports |
|--------|-------------|
| `constants` | CL, CT, CB, CW, GAP_XS/SM/MD/LG, TITLE/BODY/SMALL/KPI 字号, 全部颜色 |
| `fonts` | `set_run_font()` (STKaiti/Calibri), `add_mixed_text()`, `get_char_width()` |
| `text_layout` | `smart_textbox()`, `add_bullet_content()`, `calc_text_height()` |
| `elements` | `Card()`, `KpiStrip()`, `SectionBlock()`, `IconBadge()`, `MetricRow()`, `DividerLine()`, `AccentLine()`, `BulletDot()` |
| `tables` | `smart_table()`, `add_table()` |
| `charts` | `safe_chart_insert()`, `setup_chart_style()`, chart generators |
| `slides` | `create_cover_slide()`, `create_content_slide()`, `create_section_divider()`, `set_dark_title()`, `add_source_footer()` |
| `slide_templates` | 19 templates (T1-T20) |
| `data_driven` | `DataRegistry`, `build_deck_from_specs()`, `TEMPLATE_ROUTER` |

---

## Building Blocks

All return `bottom_y_mm` (float) for chainable layout.

| Block | 用途 |
|-------|------|
| `Card(x,y,w,h,header,body,color)` | 完整卡片组件: 2mm左accent线 + colored header bar + body panel |
| `KpiStrip(x,y,kpis)` | KPI指标行 `[("$1.2B","Revenue"), ...]` |
| `SectionBlock(x,y,w,title,color)` | 页内分区标题块: 全宽彩色方块+白色文字 |
| `ContentPanel(x,y,w,h,items)` | 文本内容面板 + smart_textbox |
| `MetricRow(x,y,metrics)` | 小指标卡行, 带趋势箭头 |
| `IconBadge(x,y,label,color)` | 彩色标签 pill (如 HIGH/MEDIUM/LOW) |
| `HeaderBar(x,y,w,h,title,color)` | 彩色标题条 |
| `DividerLine(x,y,w)` | 横向分隔线 |
| `AccentLine(x,y,h,color)` | 纵向强调线 |
| `BulletDot(x,y,color)` | 彩色圆点 marker |

---

## Slide Templates (19 templates)

| # | Function | Use |
|---|----------|-----|
| T1 | `template_kpi_dashboard` | KPI行 + bullets |
| T2 | `template_value_chain_flow` | 横向流程 + table |
| T3 | `template_chart_plus_analysis` | 图表 + 分析文字 |
| T4 | `template_comparison_matrix` | Callout + table + 结论 |
| T5 | `template_two_column_analysis` | 双栏分析 |
| T6 | `template_sidebar_case_study` | 1/4 深色侧栏 + 3/4 正文 |
| T7 | `template_three_column_compare` | 三栏对比 |
| T8 | `template_stacked_cases` | 上下两个 case |
| T9 | `template_risk_cards` | 风险卡片 |
| T10 | `template_action_plan` | 优先级漏斗 + 时间线 |
| T11 | `template_donut_chart` | 环形图 + 洞察 |
| T12 | `template_before_after` | Before/After 对比 |
| T13 | `template_funnel` | 漏斗图 |
| T14 | `template_swot` | SWOT 2×2 矩阵 |
| T15 | `template_waterfall` | 瀑布图 |
| T16 | `template_stakeholder_map` | 利益相关者地图 |
| T17 | `template_timeline` | 里程碑时间线 |
| T19 | `template_number_story` | 大数字叙事卡 |
| **T20** | **`template_executive_summary`** | **Executive Summary (lead-in + highlights + transaction)** |

---

## Section-Based Workflow

固化的工作流: 用户指定每 section 页数 → 系统按比例分配子页面。

```
用户: "做 ABC 公司 IC memo, 行业 12 页, 公司 7 页, 财务 8 页"

输出结构:
┌──────────────────────────────────────┐
│ Cover (1p)                            │
│ Executive Summary (1p) — T20          │
│ TOC (1p)                              │
├──────────────────────────────────────┤
│ [Divider] 01  行业与市场              │
│  ├─ 宏观TAM & 增速     ~18% (2p)     │
│  ├─ 市场结构 & 细分    ~18% (2p)     │
│  ├─ 竞争格局           ~18% (2p)     │
│  ├─ 行业趋势 & 驱动    ~18% (2p)     │
│  ├─ 政策与监管          ~9% (1p)     │
│  ├─ 中国视角 & 国产替代 ~18% (2p)    │
│  └─ 行业小结            ~9% (1p)     │ = 12p
├──────────────────────────────────────┤
│ [Divider] 02  公司分析                │
│  ├─ 公司概览 & 里程碑   ~28% (2p)    │
│  ├─ 商业模式 & 产品     ~28% (2p)    │
│  ├─ 竞争优势 & 护城河   ~14% (1p)    │
│  ├─ 增长策略 & pipeline ~14% (1p)    │
│  └─ 风险 & 缓释         ~14% (1p)    │ = 7p
├──────────────────────────────────────┤
│ [Divider] 03  财务与回报              │
│  ├─ 历史财务表现        ~12% (1p)    │
│  ├─ P&L 预测           ~25% (2p)    │
│  ├─ BS & CF             ~12% (1p)    │
│  ├─ Valuation           ~25% (2p)    │
│  ├─ Returns (IRR/MOIC)  ~12% (1p)    │
│  └─ 敏感性分析          ~12% (1p)    │ = 8p
├──────────────────────────────────────┤
│ Appendix (2p)                         │
└──────────────────────────────────────┘
总计: ~32p
```

### 页数分配算法

```
total_pages = section_pages
alloc = {
    "industry": {"tam":0.18, "structure":0.18, "competition":0.18, 
                  "trends":0.18, "policy":0.09, "china":0.18, "summary":0.09},
    "company":  {"overview":0.28, "business":0.28, "moat":0.14, 
                  "strategy":0.14, "risk":0.14},
    "financial":{"historical":0.12, "pnl":0.25, "bscf":0.12,
                  "valuation":0.25, "returns":0.12, "sensitivity":0.12},
}
# 每项最少 1 页, 多余页按比例分配
```

---

## Typical Creation Pattern

```python
# 1. Cover
create_cover_slide(prs, fund_name="Cathay Fund", company_name="ABC Corp", subtitle="Investment Memo")

# 2. Executive Summary
template_executive_summary(prs, lead_in="...", highlights=[...], transaction={...})

# 3. Section Divider
create_section_divider(prs, "01  行业与市场")

# 4. Content page
slide = create_content_slide(prs, topic="市场规模", conclusion="2030E $430B, CAGR 25%")
body_y = KpiStrip(slide, X1, CT, [("$430B","2030E"),("25%","CAGR")])
sec_y = SectionBlock(slide, X1, body_y + GAP_SM, CW, "驱动因素", color=CATHAY_RED)
smart_textbox(slide, X1, sec_y + GAP_XS, CW, bullet_items, max_bottom_mm=CB,
              start_font=BODY_FONT_PT, min_font=SMALL_FONT_PT)
add_source_footer(slide, "Sources...")
```

---

## Key Rules

1. 所有文本通过 `set_run_font()` 或 `add_mixed_text()` 设置字体 — 禁止直接 `run.font.name =`
2. 所有 textbox 用 `smart_textbox()` — 禁止 `slide.shapes.add_textbox()` 盲猜高度
3. 图表插入用 `safe_chart_insert()` — 自动 overflow 保护
4. 保存前 `validate_and_fix(prs)` — 自动检测文字溢出
5. 标题 = 完整结论句, ≤28字, 带数字
6. 禁止: 渐变/3D/阴影/emoji/饼图/纯黑#000/第二装饰彩
7. 禁用词: 赋能/颠覆/生态/闭环/一站式/全链路/显著/大幅
