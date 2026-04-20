# Slide Structures

# Deep Mode (25-33 slides)

参考 `ppt-generation-rules.md` 获取图表插入、字体、内容深度和布局规则。
参考 `cathay-ppt-template` skill 获取helper函数和grid常量。

**全局规则**:
- 每页内容至少200字（中文约400字），thesis/CIO页至少300字
- 图表使用 `safe_chart_insert()` 插入（width-only，返回bottom_y）
- 文本框使用 `safe_textbox()` 创建（auto-fill剩余空间）
- 中英文混合文本拆分为多个runs（`add_mixed_text()`）
- 至少使用5种不同grid pattern
- 每页必须有source footer（`add_source_footer()`）
- 所有Layout [4] slides使用 `set_title_with_conclusion(topic, conclusion)`
- 内容使用3级bullet层次（`add_bullet_content()` level 0/1/2）
- 有2+个level-0 headers的slide添加section marker icons

**Content Zone**:
- Title bar: y = 0 to ~29.2mm (CT)
- Usable content: y = 29.2mm (CT) to 175mm (CONTENT_BOTTOM_MM)
- Source footer: y = 180mm (fixed)
- Chart+Text stacking: `chart_bottom = safe_chart_insert(...); _, tf = safe_textbox(slide, x, chart_bottom+GAP_V, w)`

**Section Icon Assignment**:
| Topic Category | Icon | Color |
|---------------|------|-------|
| 财务数据/收入/利润/EPS/现金流 | ICON_FINANCIAL | Gold square |
| 投资论点/核心洞察 | ICON_INSIGHT | Red circle |
| 风险/威胁/Bear case | ICON_RISK | Red triangle |
| 催化剂/时间表 | ICON_CATALYST | Gold diamond |
| 行动计划/决策/CIO | ICON_ACTION | Red arrow |

---

## Slide 1: Cover
- **Layout**: [0] Red Title
- **Title**: N/A (cover slide — use custom cover layout)
- **Content mode**: custom
- **内容**: 公司名(中+英) + ticker + "Deep Equity Research" + 日期 + 如有upcoming earnings标注日期
- **字体**: 公司名 28pt bold white, subtitle 18pt white, 日期 14pt gold

## Slide 2: 投资摘要
- **Layout**: [4] + **KPI Row** (6格)
- **Title**: `set_title_with_conclusion("投资摘要", f"{RATING}, 目标价${TARGET} ({UPSIDE}%)")`
- **Zones**: Zone 1=KPI row (y=CT, h=28mm), Zone 2=body text (y=CT+28, h=auto-fill)
- **Content mode**: prose (thesis-style paragraphs via `add_mixed_text`)
- **Icons**: none (prose mode)
- **Grid**: `add_kpi_row()` 顶部 + full-width body 下方
- **KPI**: Price | Target | Rating | Conviction | Upside | Fwd P/E
- **Body** (≥250字): 4-5段完整thesis。每段包含：观点 + 数据证据 + 逻辑推理。不是简单bullet point，是有论证的段落。例如："HBM正在从根本上改变Micron的利润结构。FQ1'26的56.1%毛利率和FQ2'26指引的68%毛利率表明，HBM mix shift带来的结构性利润提升已经超越了传统DRAM周期波动的范畴。以FY2026E全年EPS约$35计算，当前$455的forward P/E仅12-13x，在AI半导体同行(NVDA 30x, AVGO 25x)中是绝对最低的..."
- **Source**: 数据来源 + 日期

## Slide 3: 评级与目标价
- **Layout**: [4] + **1/3 + 2/3 Grid**
- **Title**: `set_title_with_conclusion("评级与目标价", f"{RATING} Conviction {GRADE}, R:R {RR}:1")`
- **Zones**: Zone 1=sidebar (x=CL, w=ONE_THIRD), Zone 2=content (x=X2_T23, w=TWO_THIRDS)
- **Content mode**: mixed (sidebar bullets + right-side chart/table + prose)
- **Icons**: none
- **左1/3** (dark sidebar): 核心指标面板
  - Rating (大字 BUY/HOLD/SELL)
  - Conviction grade
  - Entry / Stop-loss / R:R
  - 72h保护状态 (if applicable)
- **右2/3**: Scenario分析
  - 3行scenarios表 (Bear/Base/Bull with price/probability/basis)
  - Football field chart (只指定宽度 Mm(150))
  - 行动计划段落 (100+字): "如果FQ2 beat + FQ3 guidance ≥ $20B..."
- **Source**

## Slide 4: Section Divider — 公司概览
- **Layout**: [11]
- **Title**: `set_title(slide, "公司概览")` (section divider exempt from conclusion format)
- **标题**: 白色 28pt
- **副标题**: light gold 16pt

## Slide 5: 公司概况
- **Layout**: [4] + **1/4 Dark Sidebar + 3/4 Content**
- **Title**: `set_title_with_conclusion("公司概况", f"{COMPANY}：{SECTOR} {INDUSTRY}, 市值${MCAP}B")`
- **Zones**: Zone 1=sidebar (x=CL, w=ONE_QUARTER), Zone 2=content (x=X2_Q34, w=THREE_QUARTER, h=auto-fill)
- **Content mode**: bullet-hierarchy (4 topics → level-0 headers)
- **Icons**: auto_assign — 业务描述(ICON_INSIGHT), 竞争优势(ICON_INSIGHT), 近期事件(ICON_CATALYST), 行业定位(ICON_INSIGHT)
- **左1/4** (dark red bg, white text): Key Facts
  - Sector / Industry
  - Market Cap
  - CEO
  - Employees
  - Founded
  - Exchange
- **右3/4** (≥300字):
  - 段落1: 公司是什么，核心业务描述（不是bullet point，是完整段落描述商业模式、产品组合、收入来源）
  - 段落2: 关键竞争优势和护城河（技术壁垒、客户粘性、规模经济，具体到这家公司）
  - 段落3: 近期重大事件/催化剂（最新财报表现、战略举措、产能扩张）
  - 段落4: 在行业中的定位（#几的玩家，市占率，vs竞争对手的差异化）
- **Source**

## Slide 6: 收入结构
- **Layout**: [4] + **1/2 + 1/2 Grid**
- **Title**: `set_title_with_conclusion("收入结构", f"{TOP_SEGMENT}占比{PCT}%，{TREND}")`
- **Zones**: Zone 1=left chart (x=CL, w=HALF), Zone 2=right chart (x=X2_HALF, w=HALF), Zone 3=text (y=max(chart_bottoms)+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL for level-0 headers
- **左半**: Product segmentation chart (只指定宽度 Mm(110))
- **右半**: BU breakdown chart (只指定宽度 Mm(110))
- **下方** (≥150字): 解读收入结构的含义。哪个segment是增长引擎？mix shift趋势？为什么这个收入结构对估值有意义？
- **Source**

## Slide 7: 管理层与治理
- **Layout**: [4] + **Full Width**
- **Title**: `set_title_with_conclusion("管理层与治理", f"{CEO}领导层，{KEY_TRACK_RECORD}")`
- **Zones**: Zone 1=full body (y=CT, w=CW, h=auto-fill)
- **Content mode**: prose
- **Icons**: none
- **内容** (≥200字): CEO + CFO + CTO/COO的背景、任期、track record。管理层在当前战略中的关键决策（如HBM押注、CapEx纪律、成本控制）。Insider trading信号（如有）。
- **Source**

## Slide 8: Section Divider — 行业与竞争
- **Layout**: [11]
- **Title**: `set_title(slide, "行业与竞争")`

## Slide 9: 行业格局
- **Layout**: [4] + **2/3 Text + 1/3 Chart**
- **Title**: `set_title_with_conclusion("行业格局", f"TAM ${TAM}B CAGR {CAGR}%, {COMPANY}是{POSITION}")`
- **Zones**: Zone 1=text (x=CL, w=TWO_THIRDS, h=auto-fill), Zone 2=chart (x=X2_T23, w=ONE_THIRD)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_INSIGHT for industry structure, ICON_CATALYST for structural change
- **左2/3** (≥250字):
  - 段落1: 行业规模(TAM)、增速、驱动力
  - 段落2: 行业结构（寡头/分散、进入壁垒、价值链位置）
  - 段落3: 正在发生的结构性变化（不是周期波动，是不可逆的）
  - 段落4: 这家公司在变化中是赢家还是输家，为什么
- **右1/3**: Market size chart 或 competitive position chart (只指定宽度)
- **Source**

## Slide 10: 竞争分析
- **Layout**: [4] + **Full Width Table + Commentary**
- **Title**: `set_title_with_conclusion("竞争分析", f"Fwd P/E {PE}x vs Peers中位数{MEDIAN}x")`
- **Zones**: Zone 1=table (y=CT, h=~60mm), Zone 2=commentary (y=table_bottom+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy (commentary below table)
- **Icons**: ICON_FINANCIAL
- **上方**: Comps mini-table (5-6 peers: Company, Ticker, Market Cap, Revenue, Growth, Margin, P/E)
- **下方** (≥200字): 对标分析。这家公司在peers中的相对优势/劣势。估值折价/溢价的原因。市场share trends。
- **Source**

## Slide 11: Section Divider — 财务分析
- **Layout**: [11]
- **Title**: `set_title(slide, "财务分析")`

## Slide 12: 季度收入趋势
- **Layout**: [4] + **Chart Top + Text Bottom**
- **Title**: `set_title_with_conclusion("季度收入趋势", f"FQ{X}'{YY} ${REV}B {DIRECTION}")`
- **Zones**: Zone 1=chart (y=CT, w=Mm(220) → safe_chart_insert → bottom_y), Zone 2=text (y=bottom_y+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL
- **上方**: Revenue trend chart (只指定宽度 Mm(220), y=Mm(30))
- **下方** (≥200字):
  - 收入加速/减速的驱动力分析
  - QoQ和YoY增速拆解
  - 如有guidance，对比guidance vs actual vs consensus
  - 对下一季度的预判
- **Source**

## Slide 13: 利润率扩张
- **Layout**: [4] + **Chart Top + Text Bottom**
- **Title**: `set_title_with_conclusion("利润率扩张", f"GM {GM}% ({STRUCTURAL_OR_CYCLICAL})")`
- **Zones**: same as Slide 12 (chart top + text bottom)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL
- **上方**: Margin trends chart (只指定宽度 Mm(220))
- **下方** (≥200字):
  - GM扩张的结构性驱动 vs 周期性因素
  - OM杠杆效应分析 (R&D/SGA占比变化)
  - 与行业标准的对比
  - 对利润率可持续性的判断
- **Source**

## Slide 14: EPS与现金流
- **Layout**: [4] + **1/2 + 1/2 Grid**
- **Title**: `set_title_with_conclusion("EPS与现金流", f"EPS ${EPS} ({YOY}% YoY), FCF ${FCF}B")`
- **Zones**: Zone 1=left chart, Zone 2=right chart, Zone 3=text (y=max(bottoms)+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL
- **左半**: EPS trend chart (只指定宽度 Mm(110))
- **右半**: Cash flow chart (只指定宽度 Mm(110))
- **下方** (≥150字):
  - EPS加速背后的驱动力（收入 vs 利润率 vs 股数）
  - FCF趋势和CapEx intensity分析
  - 现金部署策略（再投资 vs 股东回报）
- **Source**

## Slide 15: 资产负债表
- **Layout**: [4] + **Full Width**
- **Title**: `set_title_with_conclusion("资产负债表", f"Net Debt/EBITDA {RATIO}x, {LEVERAGE_TREND}")`
- **Zones**: Zone 1=table (y=CT), Zone 2=commentary (y=table_bottom+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL
- **上方**: BS summary table (Assets/Debt/Cash/Equity, 2 quarters对比)
- **下方** (≥150字):
  - 杠杆水平和趋势（去杠杆/加杠杆）
  - 流动性状况（current ratio, cash coverage）
  - 库存健康度（DIO趋势，是否有积压风险）
  - Working capital efficiency
- **Source**

## Slide 16: Section Divider — 估值分析
- **Layout**: [11]
- **Title**: `set_title(slide, "估值分析")`

## Slide 17: 估值对标
- **Layout**: [4] + **1/3 Metrics + 2/3 Chart**
- **Title**: `set_title_with_conclusion("估值对标", f"Fwd P/E {PE}x, Peers中最{POSITION}")`
- **Zones**: Zone 1=metrics sidebar (x=CL, w=ONE_THIRD), Zone 2=chart (x=X2_T23, w=TWO_THIRDS), Zone 3=text below
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL
- **左1/3**: Key valuation metrics面板
  - Forward P/E
  - EV/EBITDA
  - P/B
  - FCF Yield
  - 与Peers中位数的对比
- **右2/3**: Peer valuation comparison chart (只指定宽度 Mm(150))
- **下方** (≥150字): 估值折价/溢价的原因分析。为什么市场给这个估值？这个估值合不合理？
- **Source**

## Slide 18: DCF分析
- **Layout**: [4] + **Chart + Commentary**
- **Title**: `set_title_with_conclusion("DCF分析", f"Fair Value ${FV} (WACC {WACC}%)")`
- **Zones**: Zone 1=chart (y=CT → safe_chart_insert → bottom_y), Zone 2=text (y=bottom_y+GAP_V, h=auto-fill)
- **Content mode**: mixed
- **Icons**: ICON_FINANCIAL
- **上方**: DCF sensitivity heatmap (只指定宽度 Mm(180))
- **下方** (≥200字):
  - WACC计算过程（Ke, Kd, weights, beta, risk premium）
  - Terminal growth rate选择的理由
  - FCF forecast的关键假设
  - 与market price的对比，implied growth rate
- **Source**

## Slide 19: Football Field
- **Layout**: [4] + **Chart + Summary Table**
- **Title**: `set_title_with_conclusion("综合估值", f"Range ${LOW}-${HIGH}, Target ${TARGET}")`
- **Zones**: Zone 1=chart (y=CT → safe_chart_insert → bottom_y), Zone 2=table+text (y=bottom_y+GAP_V, h=auto-fill)
- **Content mode**: mixed
- **Icons**: none
- **上方**: Football field chart (只指定宽度 Mm(220))
- **下方**: 方法对比表 (Method | Low | High | Midpoint)
- **Commentary** (≥100字): 综合各方法的结论，目标价的选择逻辑
- **Source**

## Slide 20: Section Divider — 投资论点
- **Layout**: [11]
- **Title**: `set_title(slide, "投资论点")`

## Slide 21: Bull Case
- **Layout**: [4] + **Full Width**
- **Title**: `set_title_with_conclusion("Bull Case", f"{ONE_LINE_BULL_THESIS}")`
- **Zones**: Zone 1=full body (y=CT, w=CW, h=auto-fill)
- **Content mode**: prose (thesis-style, use add_mixed_text for paragraphs)
- **Icons**: none (prose mode)
- **内容** (≥300字):
  - 段落1: 核心做多论点（3句话thesis完整展开为一段话）
  - 段落2: 催化剂时间表（不只列表，要分析每个催化剂的逻辑链）
  - 段落3: 上行情景量化（Base/Bull/Super Bull，每个有计算过程）
  - 段落4: 预判Bear反驳并预先回应
  - Conviction Score + 理由
- **Source**

## Slide 22: Bear Case
- **Layout**: [4] + **Full Width**
- **Title**: `set_title_with_conclusion("Bear Case", f"{ONE_LINE_BEAR_THESIS}")`
- **Zones**: Zone 1=full body (y=CT, w=CW, h=auto-fill)
- **Content mode**: prose
- **Icons**: none
- **内容** (≥300字):
  - 段落1: 核心做空论点（3句话完整展开）
  - 段落2: Kill Score + 证据链（不只给分数，要展开每条证据）
  - 段落3: 逆向估值压力测试（70%/50%/30%情景分析，有计算）
  - 段落4: Top 3致命风险的详细分析
  - 段落5: 历史类比
- **Source**

## Slide 23: CIO裁决
- **Layout**: [4] + **1/4 Dark Sidebar + 3/4 Content**
- **Title**: `set_title_with_conclusion("CIO裁决", f"{DECISION} Conviction {GRADE}, Watch {KEY_ITEM}")`
- **Zones**: Zone 1=sidebar (x=CL, w=ONE_QUARTER), Zone 2=content (x=X2_Q34, w=THREE_QUARTER, h=auto-fill)
- **Content mode**: prose (CIO reasoning requires full paragraphs)
- **Icons**: none
- **左1/4** (dark sidebar, white text):
  - Decision: BUY/HOLD/SELL (大字)
  - Conviction: A/B/C/D
  - Target: $XXX
  - Stop: $XXX
  - R:R: X.X:1
- **右3/4** (≥350字):
  - 段落1: CIO独立研判 — 生意本质判断（这不是复述Bull/Bear，是独立观点）
  - 段落2: 估值框架判断 — 市场用什么框架定价，对不对，会不会切换
  - 段落3: Elephant in the Room — 所有分析师忽略的最大风险或机会
  - 段落4: 论证质量评估 — Bull/Bear哪边更强，为什么
  - 段落5: 最终决策理由 + 行动计划
- **Source**

## Slide 24: Section Divider — 风险与监控
- **Layout**: [11]
- **Title**: `set_title(slide, "风险与监控")`

## Slide 25: 风险矩阵
- **Layout**: [4] + **Full Width Table + Commentary**
- **Title**: `set_title_with_conclusion("风险矩阵", f"Top Risk: {RISK_1}")`
- **Zones**: Zone 1=table (y=CT), Zone 2=commentary (y=table_bottom+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy (commentary uses level-0 + level-1)
- **Icons**: ICON_RISK
- **上方**: Risk table (Risk | Probability | Impact | Evidence | Mitigation)
  - 至少5个风险，覆盖: Company-specific, Industry, Financial, Macro
- **下方** (≥150字): 风险之间的关联性分析。哪些风险是相关的？什么情况下会同时发生？最需要关注的1-2个风险是什么？
- **Source**

## Slide 26: 监控计划
- **Layout**: [4] + **1/3 + 1/3 + 1/3 三列**
- **Title**: `set_title_with_conclusion("监控计划", f"Next: {CATALYST} ({DATE})")`
- **Zones**: 3 equal columns (x=CL/X2_MID/X3_RIGHT, w=THIRD each)
- **Content mode**: bullet-hierarchy (each column has level-0 header)
- **Icons**: ICON_CATALYST for each column header
- **Column 1**: 短期监控 (7天)
  - 即将到来的催化剂
  - 财报行动计划（if applicable）
  - 具体的buy/sell trigger
- **Column 2**: 中期监控 (30天)
  - 领先恶化指标（3-5个可观测信号）
  - 升级/降级触发条件
  - 止损规则
- **Column 3**: 长期关注 (3-6个月)
  - 行业结构变化追踪
  - 叙事完整性评估
  - 竞争对手动态
- 每列至少100字
- **Source**

## Slide 27: Section Divider — 附录
- **Layout**: [11]
- **Title**: `set_title(slide, "附录")`

## Slides 28-31: MoA辩论完整记录 (Appendix)
- **Layout**: [4] + **Full Width**
- **Title**: `set_title_with_conclusion("附录: {AGENT_NAME}", f"{KEY_CONCLUSION}")`
- **Zones**: Zone 1=full body (y=CT, w=CW, h=auto-fill)
- **Content mode**: mixed (prose paragraphs with key data)
- **Icons**: none
- **Slide 28**: Phase 1 Foundation — 宏观策略师 + 法务会计师 + 行业专家的完整输出 (≥300字)
- **Slide 29**: Bull Advocate 完整论证 — thesis, catalysts表, price targets, 预判反驳 (≥300字)
- **Slide 30**: Bear Advocate 完整论证 — short thesis, kill score证据, 逆向估值, top risks (≥300字)
- **Slide 31**: Rebuttal Exchange + CIO评分表 — bull/bear互相反驳的要点, CIO对每个论点的评分 (≥250字)

## Slide 32: 数据来源
- **Layout**: [4] + **Full Width**
- **Title**: `set_title(slide, "数据来源与方法论")` (exempt from conclusion format)
- **Content mode**: bullet-hierarchy (level-0 for each source category)
- 所有数据来源列表 + 日期
- API来源: FMP, Seeking Alpha, Brave Search
- 内部来源: Knowledge Base, Worldview, Position Thesis
- 免责声明

## Slide 33: 免责声明
- **Layout**: [4]
- **Title**: `set_title(slide, "免责声明")` (exempt)
- **Content mode**: prose (8pt grey)
- Standard disclaimer (8-9pt grey text)

---

## Chart Specifications

所有图表使用cathay-ppt-template的matplotlib helpers:
- `cathay_bar_chart()`, `cathay_line_chart()`, `cathay_pie_chart()`, `cathay_waterfall_chart()`
- 颜色: `CATHAY_COLORS = ['#800000', '#C8A415', '#808080', '#E60000', '#E8D590', '#404040']`
- DPI: 200 (cathay-ppt-template默认)
- 背景: white
- Grid: y-axis only, alpha=0.3

**Deep Mode图表清单** (15-20张):

| # | Chart | figsize | 插入宽度 | 所在Slide |
|---|-------|---------|---------|----------|
| 1 | Revenue trend (quarterly) | (8, 4.5) | Mm(220) | 12 |
| 2 | Margin trends (GM/OM) | (8, 4.5) | Mm(220) | 13 |
| 3 | EPS trend (quarterly) | (8, 4.5) | Mm(110) | 14 |
| 4 | Cash flow (OCF/CapEx/FCF) | (8, 4.5) | Mm(110) | 14 |
| 5 | Revenue by product | (6, 4.5) | Mm(110) | 6 |
| 6 | Revenue by BU | (7, 4.5) | Mm(110) | 6 |
| 7 | Peer valuation (fwd P/E) | (7, 4.5) | Mm(150) | 17 |
| 8 | DCF sensitivity heatmap | (7, 5) | Mm(180) | 18 |
| 9 | Football field | (8, 4) | Mm(220) | 19 |
| 10 | Scenario comparison | (7, 4.5) | Mm(150) | 3 |
| 11 | Market size/share | (6, 4.5) | Mm(100) | 9 |
| 12 | Competitive positioning | (6, 4.5) | Mm(100) | 10 |

---

# Slide Structures (Quick Mode)

Quick模式产出10-12页精简报告，保留核心分析框架但合并相关内容。

**全局规则** (与Deep Mode一致):
- 每页内容至少200字（中文约400字），thesis页至少250字
- 图表使用 `safe_chart_insert()` 插入（width-only，返回bottom_y）
- 文本框使用 `safe_textbox()` 创建（auto-fill剩余空间）
- 中英文混合文本拆分为多个runs（`add_mixed_text()`）
- 至少使用4种不同grid pattern
- 每页必须有source footer（`add_source_footer()`）
- 所有Layout [4] slides使用 `set_title_with_conclusion(topic, conclusion)`
- 内容使用3级bullet层次（`add_bullet_content()` level 0/1/2）
- 有2+个level-0 headers的slide添加section marker icons

---

## QSlide 1: Cover
- **Layout**: [0] Red Title
- **Title**: N/A (cover slide — use custom cover layout)
- **Content mode**: custom
- **内容**: 公司名(中+英) + ticker + "Equity Research — Quick Screening" + 日期 + 如有upcoming earnings标注日期
- **字体**: 公司名 28pt bold white, subtitle 18pt white, 日期 14pt gold

## QSlide 2: 投资摘要
- **Layout**: [4] + **KPI Row** (6格)
- **Title**: `set_title_with_conclusion("投资摘要", f"{RATING}, 目标价${TARGET}")`
- **Zones**: Zone 1=KPI row (y=CT, h=28mm), Zone 2=body text (y=CT+28, h=auto-fill)
- **Content mode**: prose (thesis-style paragraphs via `add_mixed_text`)
- **Icons**: none (prose mode)
- **Grid**: `add_kpi_row()` 顶部 + full-width body 下方
- **KPI**: Price | Target | Rating | Upside | Fwd P/E | Market Cap
- **Body** (≥250字): 3-4段完整thesis。每段：观点 + 数据证据 + 逻辑推理。包含：核心投资逻辑、估值判断、关键催化剂。
- **Source**

## QSlide 3: 公司概况
- **Layout**: [4] + **1/4 Dark Sidebar + 3/4 Content**
- **Title**: `set_title_with_conclusion("公司概况", f"{COMPANY}：{SECTOR}, 市值${MCAP}B")`
- **Zones**: Zone 1=sidebar (x=CL, w=ONE_QUARTER), Zone 2=content (x=X2_Q34, w=THREE_QUARTER, h=auto-fill)
- **Content mode**: bullet-hierarchy (3 topics → level-0 headers)
- **Icons**: ICON_INSIGHT
- **左1/4** (dark red bg, white text): Key Facts
  - Sector / Industry
  - Market Cap
  - CEO
  - Employees
  - Exchange
- **右3/4** (≥250字):
  - 段落1: 商业模式、核心产品/服务、收入来源
  - 段落2: 竞争优势和护城河
  - 段落3: 近期重大事件/催化剂
- **Source**

## QSlide 4: 收入结构与分拆
- **Layout**: [4] + **1/2 + 1/2 Grid**
- **Title**: `set_title_with_conclusion("收入结构", f"{TOP_SEGMENT}占比{PCT}%，{TREND}")`
- **Zones**: Zone 1=left chart (x=CL, w=HALF), Zone 2=right chart (x=X2_HALF, w=HALF), Zone 3=text (y=max(chart_bottoms)+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL for level-0 headers
- **左半**: Product segmentation chart (只指定宽度 Mm(110))
- **右半**: Revenue trend chart (只指定宽度 Mm(110))
- **下方** (≥150字): 收入结构含义、增长引擎segment、mix shift趋势
- **Source**

## QSlide 5: 季度财务趋势
- **Layout**: [4] + **Chart Top + Text Bottom**
- **Title**: `set_title_with_conclusion("季度财务趋势", f"GM {GM}% {TREND}")`
- **Zones**: Zone 1=chart (y=CT, w=Mm(220) → safe_chart_insert → bottom_y), Zone 2=text (y=bottom_y+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL
- **上方**: Margin trends chart (只指定宽度 Mm(220))
- **下方** (≥200字):
  - 季度收入加速/减速驱动力
  - GM/OM趋势及结构性 vs 周期性判断
  - Guidance vs actual vs consensus对比（如有）
  - 对下一季度的预判
- **Source**

## QSlide 6: EPS与现金流
- **Layout**: [4] + **1/2 + 1/2 Grid**
- **Title**: `set_title_with_conclusion("EPS与现金流", f"EPS ${EPS} ({YOY}% YoY), FCF ${FCF}B")`
- **Zones**: Zone 1=left chart, Zone 2=right chart, Zone 3=text (y=max(bottoms)+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL
- **左半**: EPS trend chart (只指定宽度 Mm(110))
- **右半**: Cash flow chart (只指定宽度 Mm(110))
- **下方** (≥150字): EPS驱动力拆解、FCF趋势、现金部署策略
- **Source**

## QSlide 7: 竞争与估值
- **Layout**: [4] + **Full Width Table + Chart**
- **Title**: `set_title_with_conclusion("竞争与估值", f"Fwd P/E {PE}x vs Peers中位数{MEDIAN}x")`
- **Zones**: Zone 1=table (y=CT, h=~55mm), Zone 2=chart (x=CL, w=HALF → safe_chart_insert → bottom_y), Zone 3=text (x=X2_HALF, w=HALF, y=table_bottom+GAP_V, h=auto-fill)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_FINANCIAL
- **上方**: Comps mini-table (5-6 peers: Company, Ticker, Mkt Cap, Revenue, Growth, P/E)
- **下方左**: Peer valuation chart (只指定宽度 Mm(110))
- **下方右** (≥150字): 相对估值定位、折价/溢价原因、top-2 paradigm估值方法结论
- **Source**

## QSlide 8: Football Field 估值
- **Layout**: [4] + **Chart + Summary Table**
- **Title**: `set_title_with_conclusion("综合估值", f"Target ${TARGET}")`
- **Zones**: Zone 1=chart (y=CT → safe_chart_insert → bottom_y), Zone 2=table+text (y=bottom_y+GAP_V, h=auto-fill)
- **Content mode**: mixed (table + commentary prose)
- **Icons**: none
- **上方**: Football field chart (只指定宽度 Mm(220))
- **下方**: 方法对比表 (Method | Low | High | Midpoint)
- **Commentary** (≥100字): 综合各方法结论，目标价选择逻辑
- **Source**

## QSlide 9: Bull / Bear Case
- **Layout**: [4] + **1/2 + 1/2 Grid**
- **Title**: `set_title_with_conclusion("Bull / Bear Case", f"Bull: {THESIS} | Bear: {THESIS}")`
- **Zones**: Zone 1=left half (x=CL, w=HALF, y=CT, h=auto-fill), Zone 2=right half (x=X2_HALF, w=HALF, y=CT, h=auto-fill)
- **Content mode**: bullet-hierarchy (each half has level-0 header)
- **Icons**: left=ICON_CATALYST (bull), right=ICON_RISK (bear)
- **左半 — Bull Case** (≥200字):
  - 核心做多论点 (thesis展开)
  - Top 3催化剂 (含日期)
  - Bull目标价 + 上行空间
- **右半 — Bear Case** (≥200字):
  - 核心做空论点 (thesis展开)
  - Top 3风险 (含概率/影响)
  - Bear目标价 + 下行空间
- **Source**

## QSlide 10: 风险与监控
- **Layout**: [4] + **2/3 Table + 1/3 Monitoring**
- **Title**: `set_title_with_conclusion("风险与监控", f"Top Risk: {RISK_1}, Next: {CATALYST} ({DATE})")`
- **Zones**: Zone 1=table (x=CL, w=TWO_THIRDS, y=CT, h=auto-fill), Zone 2=monitoring (x=X2_T23, w=ONE_THIRD, y=CT, h=auto-fill)
- **Content mode**: bullet-hierarchy
- **Icons**: ICON_RISK for risk headers, ICON_CATALYST for monitoring headers
- **左2/3**: Risk table (Risk | Probability | Impact | Mitigation) — 至少5个风险
- **右1/3**: 监控计划
  - 短期 (7天): buy/sell trigger、即将催化剂
  - 中期 (30天): 恶化指标、升降级条件
- **Source**

## QSlide 11: 数据来源
- **Layout**: [4] + **Full Width**
- **Title**: `set_title(slide, "数据来源")` (exempt from conclusion format)
- **Zones**: Zone 1=full body (y=CT, w=CW, h=auto-fill)
- **Content mode**: bullet-hierarchy (level-0 for each source category)
- **Icons**: none
- 所有数据来源列表 + 日期
- 免责声明

## QSlide 12 (可选): 附加分析
- **Layout**: [4] + **Full Width**
- **Title**: `set_title_with_conclusion("附加分析", f"{TOPIC_SUMMARY}")` (or `set_title(slide, "附加分析")` if no clear conclusion)
- **Zones**: Zone 1=full body (y=CT, w=CW, h=auto-fill)
- **Content mode**: mixed (adapt to the specific additional content)
- **Icons**: assign per content topic (ICON_INSIGHT / ICON_CATALYST / ICON_FINANCIAL as appropriate)
- 如有额外内容（持仓状态详情、earnings preview视角、行业特殊分析），放在此页
- 无附加内容时省略此页

---

## Quick Mode图表清单 (5-8张)

| # | Chart | figsize | 插入宽度 | 所在QSlide |
|---|-------|---------|---------|-----------|
| 1 | Revenue by product/segment | (6, 4.5) | Mm(110) | 4 |
| 2 | Revenue trend (quarterly) | (8, 4.5) | Mm(110) | 4 |
| 3 | Margin trends (GM/OM quarterly) | (8, 4.5) | Mm(220) | 5 |
| 4 | EPS trend (quarterly) | (8, 4.5) | Mm(110) | 6 |
| 5 | Cash flow (OCF/CapEx/FCF) | (8, 4.5) | Mm(110) | 6 |
| 6 | Peer valuation comparison | (7, 4.5) | Mm(110) | 7 |
| 7 | Football field | (8, 4) | Mm(220) | 8 |
| 8 | Scenario comparison (可选) | (7, 4.5) | Mm(150) | 2 |
