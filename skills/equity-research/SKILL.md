---
name: equity-research
description: Use when asked to create equity research reports, stock analysis decks, investment analysis PPT, or when user invokes /equity-research. Supports quick (10-15 slides) and deep (25-40 slides with MoA debate) modes. Produces Cathay-branded PPT + PDF, optionally Excel.
---

# Equity Research

## Overview

Unified equity research skill that produces Cathay-branded PPT + PDF (always) and Excel financial model (deep mode only). Two modes: `quick` for rapid screening and `deep` for full MoA debate analysis.

## Invocation

```
/equity-research {TICKER} quick    → 10-15 slides, ~5-10 min
/equity-research {TICKER} deep     → 25-40 slides + Excel, ~30-60 min
```

If mode not specified, ask user. Default to `quick` for time-sensitive requests.

**Coverage**: US stocks (primary, full data), HK/A-shares (secondary, web + KB).

## Output

| Deliverable | Quick | Deep |
|-------------|-------|------|
| PPT (.pptx) | Yes | Yes |
| PDF (.pdf) | Yes | Yes |
| Excel (.xlsx) | No | Yes (3 sheets) |

File naming: `{TICKER}_Equity_Research_{YYYYMMDD}.pptx/.pdf`, `{TICKER}_Financial_Model_{YYYYMMDD}.xlsx`

---

## Quick Mode Workflow

### Phase 1: Data Collection

**即使是Quick模式也必须遵守数据完整性铁律** (见Deep Mode Phase 1)。区别仅在于拉取范围较小。

1. **Identify company**: FMP `profile` → sector, industry, market cap
2. **Quarterly financials**: FMP `income-statement` period=quarter limit=6 → **季度趋势**（不只看年度）
3. **Earnings date check**: 确认下一个财报日期，如7天内则标注
4. **Street consensus**: Brave search + SA price target → 至少2个交叉验证源
5. **Position check**: 读 position_thesis.md 确认是否活跃持仓
6. **News**: Brave search latest news
7. **Paradigm routing**: Match sector/industry → paradigm → select top-2 valuation methods

### Phase 2: Analysis

Single-pass fundamental analysis:
1. Company overview (business model, products, market position)
2. Financial highlights (revenue/margin trends, key metrics)
3. Competitive positioning (top 3-5 peers, moat assessment)
4. Valuation (apply top-2 paradigm methods)
5. Investment conclusion: **BUY/HOLD/SELL** + target price + upside %
6. Bull/bear case (3-4 bullets each side)
7. Top 5 risks (probability + impact)
8. Key catalysts with dates

### Phase 3: Report Generation

1. **Charts**: Generate 5-8 charts via matplotlib with Cathay palette
   - **REQUIRED**: Use `cathay-ppt-template` chart helpers (`cathay_bar_chart`, `cathay_line_chart`, `cathay_pie_chart`)
   - See chart list in `references/slide-structures.md` Quick Mode section
2. **PPT**: Build slides per `references/slide-structures.md` Quick Mode sequence
   - **REQUIRED**: Use `cathay-ppt-template` for all slide creation (layouts, text formatting, table styling)
   - Template: `~/.claude/skills/cathay-ppt-template/assets/template.pptx`
3. **QC + Export**:
   - `validate_no_overlap()` — check for shape overlaps
   - `export_to_pdf()` — generate PDF via LibreOffice
   - Verify all slides have source footers

---

## Deep Mode Workflow

### Phase 1: Data Collection (Enhanced)

**DATA INTEGRITY RULES (铁律)**:
1. **季度数据优先于年度数据** — 年度汇总掩盖趋势拐点，必须拉最近6-8个季度的IS/BS/CF
2. **必须确认下一个财报日期** — 用 FMP `earning-calendar` 或 Brave 搜索确认。如果财报在7天内，分析必须以earnings preview视角展开
3. **Street consensus 必须交叉验证** — FMP analyst-estimates + Seeking Alpha (RapidAPI) + Brave搜索至少2个源。如果API返回空值，用Brave搜索补充
4. **公司指引 vs Street共识必须分开标注** — 不能混淆management guidance和analyst consensus
5. **检查是否为活跃持仓** — 读取 `~/Algo Trading/Quant Trading/机长进化论/memory/shared/position_thesis.md`，如有持仓则标注入场价/Kill Score/止损位
6. **KB Deep 必须检查** — 搜索 `~/Algo Trading/Quant Trading/机长进化论/memory/kb_deep/` 和 `~/.openclaw/workspace/memory/` 中的相关材料
7. **Worldview 必须读取** — 读取 `~/Algo Trading/Quant Trading/机长进化论/memory/shared/synthesis/worldview.md` 获取当日宏观regime和信号聚类

**数据拉取清单** (全部执行，不可跳过):
1. FMP `profile` → sector, industry, market cap
2. FMP `income-statement` period=quarter limit=8 → **季度趋势是核心**
3. FMP `income-statement` period=annual limit=5 → 年度汇总
4. FMP `balance-sheet-statement` period=quarter limit=2 → 最新BS
5. FMP `cash-flow-statement` period=quarter limit=8 → 季度现金流趋势
6. FMP `revenue-product-segmentation` period=quarter → 收入分拆
7. FMP `stock-price-change` → 价格表现
8. FMP `earnings-surprises` → beat/miss历史
9. Seeking Alpha `symbols/get-analyst-price-target` → 分析师目标价
10. Seeking Alpha `symbols/get-summary` → Forward PE等
11. Brave search: `"{COMPANY} fiscal Q{X} {YEAR} earnings consensus estimate"` → Street共识
12. Brave search: `"{COMPANY} guidance outlook revenue"` → 公司指引
13. Brave search: `"{COMPANY} {INDUSTRY} recent news"` → 最新新闻
14. Position thesis check → 是否为活跃持仓
15. Worldview read → 宏观regime + 今日信号
16. KB Deep search → 相关深度材料

**CHECKPOINT 1**: Present data summary including quarterly trajectory, upcoming earnings date, Street consensus vs guidance, position status. Wait for user confirmation.

### Phase 2: MoA 7-Agent Debate

Follow `references/moa-debate-framework.md` exactly. Execute each agent role sequentially:

**Phase 1 — Foundation** (label each output clearly):
1. **[Macro Strategist]**: Economic environment, rates, sector momentum
2. **[Forensic Accountant]**: Financial quality, cash flow, red flags → quality score (1-10)
3. **[Industry Specialist]**: Use paradigm prompt fragment, competitive landscape, moat → advantage assessment

**Phase 2 — Advocacy**:
4. **[Bull Advocate]**: Full bull case — thesis, catalysts (with dates), 3 price targets, pre-emptive bear rebuttals, conviction score (1-100)
5. **[Bear Advocate]**: Full bear case — short thesis, kill score (1-10), reverse valuation stress test, top 3 risks, historical analogy, deterioration indicators

**Phase 2.5 — Rebuttals**:
6. **[Bull Rebuttal]**: Respond to Bear's top 3 arguments
7. **[Bear Rebuttal]**: Respond to Bull's top 3 arguments

**Phase 3 — CIO Decision**:
8. **[CIO]**:
   - Step 1: Independent view FIRST (business essence, industry structure, valuation framework judgment, elephant in room) — form this BEFORE reading Bull/Bear
   - Step 2: Evaluate Bull/Bear argument quality (score each point 1-10)
   - Step 3: Blind spots
   - Step 4: Decision (STRONG_BUY/BUY/HOLD/SELL/STRONG_SELL/NO_POSITION)
   - Step 5: Conviction grade (A/B/C/D), position sizing, price targets, stop-loss

### Phase 3: Valuation

Apply ALL methods from matched paradigm (3-4 methods):
- DCF with proper WACC (cost of equity, cost of debt, weights)
- Comps analysis (8-10 peers with statistical summary)
- Paradigm-specific methods (NAV, P/NAV, FCF Yield, etc.)
- Cross-validate with CIO price targets
- Build football field chart data
- Build sensitivity tables (WACC vs terminal growth)

**CHECKPOINT 2**: Present to user:
```
CIO 决策: {DECISION} — Conviction {GRADE}
目标价: ${TARGET} (upside +XX%)
止损: ${STOP_LOSS}
R:R = X.X:1

估值摘要:
■ DCF: $XXX (WACC: X.X%, Terminal Growth: X.X%)
■ {Method 2}: $XXX
■ {Method 3}: $XXX
■ Football Field: $XXX - $XXX

关键辩论要点:
■ Bull 最强论点: {strongest_bull_point}
■ Bear 最强论点: {strongest_bear_point}
■ CIO 独立洞察: {cio_original_insight}

是否继续生成报告？
```

### Phase 4: Report Generation

**MUST READ**: `references/ppt-generation-rules.md` 在生成任何PPT代码之前。违反规则 = 低质量输出。

**PPT Generation Iron Rules (写代码前必须遵守)**:
1. **Chart只指定宽度** — `add_picture(path, Mm(x), Mm(y), Mm(w))`，不传height参数
2. **字体必须用set_run_font()** — 完整XML属性版本，不是简化版。混合中英文必须拆分runs
3. **每页≥200字** — 不允许"3个bullet"式浅薄slide。写段落，不只是列表
4. **布局多样化** — 至少5种不同grid pattern，不允许连续3页相同layout
5. **必须用cathay-ppt-template helpers** — setup_text_frame, format_paragraph, set_run_font, set_square_bullet, add_bullet_content, add_source_footer, add_table, add_kpi_row
6. **必须用结论式标题** — `set_title_with_conclusion(topic, conclusion)`，不是 `set_title()`

**生成步骤**:
1. **Charts**: 生成12-20张图表，保存为PNG (见 `references/slide-structures.md` chart清单)
2. **PPT**: 按 `references/slide-structures.md` 逐页构建:
   - 图表用 `safe_chart_insert()` 插入（返回bottom_y）
   - 文本框用 `safe_textbox()` 创建（auto-fill剩余空间）
   - 标题用 `set_title_with_conclusion(topic, conclusion)`
   - 内容用 `add_bullet_content()` 3级层次
   - 有2+ level-0 headers的slide添加 `add_section_marker()` icons
3. **Excel**: 3 sheets with `cathay-excel-template` branding
4. **QC**: 运行 `qc_presentation()` (包含 `validate_no_overlap()` + `validate_text_fit()` + `export_to_pdf()`) + 检查字体/字数/布局多样性/结论标题/icon

---

## Language & Typography

- Chinese body text → **楷体** (KaiTi)
- English text and numbers → **Calibri**
- Use `set_run_font()` from `cathay-ppt-template` for auto-detection per text run
- Slide titles: 中文 (e.g., "投资摘要", "估值分析", "风险与监控")
- Data labels, axis labels: Calibri
- Source footers: 8pt, grey (#808080), `add_source_footer()`
- All MoA analysis output: 简体中文

---

## Quality Checklist

Before delivering output, verify:

**图表质量**:
- [ ] 所有图表插入只指定宽度，不指定高度（不变形）
- [ ] Charts use Cathay color palette (#800000, #C8A415, #E8D590)
- [ ] 所有图表通过 `safe_chart_insert()` 插入（不直接用 `add_picture`）

**字体**:
- [ ] 中文文本 → 楷体 (通过set_run_font XML属性设置)
- [ ] 英文/数字 → Calibri
- [ ] 混合中英文文本已拆分为独立runs

**内容深度**:
- [ ] 每个内容slide至少200字（中文约400字）
- [ ] Thesis/CIO/Bull/Bear slides至少300字
- [ ] 内容是段落式分析，不只是bullet point列表
- [ ] 每段有: 观点 → 数据支撑 → 推理/含义
- [ ] 内容slides使用level-0/1/2 bullet层次（3+ topics有level-0 headers）
- [ ] 有2+ level-0 headers的slide有section marker icons

**布局多样性**:
- [ ] 至少使用5种不同grid pattern
- [ ] 没有连续3页以上使用相同layout
- [ ] 使用了KPI row, dark sidebar, split columns, full-width table等多种模式

**格式与输出**:
- [ ] All slides have source footer (via `add_source_footer()`)
- [ ] Tables use Cathay formatting (red header, white text, alternating rows)
- [ ] No shape overlaps (`validate_no_overlap()`)
- [ ] PDF renders correctly (`export_to_pdf()`)
- [ ] Investment Summary has: rating + target price + upside %
- [ ] Football field chart present
- [ ] *Deep mode*: PPT and Excel numbers are consistent
- [ ] *Deep mode*: Full MoA debate in Appendix slides (each ≥250字)
- [ ] 所有Layout [4] slides使用 `set_title_with_conclusion()`（标题含核心结论）
- [ ] `validate_text_fit()` 返回零warnings
- [ ] 文本框通过 `safe_textbox()` 创建（无hardcoded过小高度）
- [ ] 无直接 `.font.name =` 设置（全部通过 `set_run_font()`）

---

## Cross-References

- **REQUIRED**: `cathay-ppt-template` — all PPT creation (layouts, helpers, charts, PDF export)
- **REQUIRED (deep)**: `cathay-excel-template` — Excel branding and formatting
- **MUST READ BEFORE PPT**: `~/.claude/skills/cathay-ppt-template/references/ppt-generation-rules.md` — 图表/字体/深度/布局/overflow/bullet/title铁律（已移至全局位置）
- `references/slide-structures.md` — slide-by-slide specs with grid pattern and word count requirements
- `references/moa-debate-framework.md` — full MoA debate prompts (deep mode only)
- `references/valuation-paradigms.md` — 9 industry paradigms with routing logic (incl. Financials, Healthcare, Consumer Staples)
- `references/data-collection-guide.md` — FMP, SA, Tavily, Brave, SEC, FRED, Tushare data sources
