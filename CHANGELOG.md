# Changelog

本项目所有重要更新记录在这里。日期按 `YYYY-MM-DD`。

---

## [2026-04-23] — Excel 模型：IB Deal-Tested Patterns

**背景**：基于深海智人 / Project Sealien 实盘建模（2026 年 4 月）对比我第一版交付与最终 banker 版本的差异，把 7 条真实踩坑后修正的 patterns 固化进 skill。Template 本身是对的；这些是 **modeling judgment** 层面的 convention。

### 新增

- `templates/cathay-excel/references/ib-deal-patterns.md` — 7 条 IB 级建模 patterns，每条带公式示例 + 对错对比 + quick-check checklist。
- `templates/cathay-excel/SKILL.md` — 新增指向 `references/ib-deal-patterns.md` 的 pointer（ship 前必读）。

### 7 条 Patterns 摘要

| # | Pattern | 关键修正 |
|---|---------|---------|
| 1 | **CFF = 本轮全额，不是 sponsor 的份** | 初创融资入账按全轮金额计入 CFF，sponsor 份额只用于稀释计算 |
| 2 | **市场规模在 P&L 顶部** | Global TAM / China SAM / 市占率放 Revenue 上方 4 行，不是 NI 下方 |
| 3 | **不要 row padding** | 删除 `while r < 10: r += 1` 式的行号填充；用 row ref 变量，Financials 压缩 35% (162→105 行) |
| 4 | **分隔 sheet `>>`** | 在 input / calc / appendix 之间插入空白 sheet `>>`，banker 习惯 |
| 5 | **Return block 行序** | NI → Hold → PE → ExitEq → Ownership → Cathay → Invested → 空行 → MOIC → IRR |
| 6 | **标题极简** | `{Company} — 财务模型 ({YYYY-MM})`，不加机构前缀 |
| 7 | **长周期工资通胀** | 5-7 年预测里管理/研发人均工资按 3-5% p.a. 增长，flat 会高估经营杠杆 |

### 怎么用

下次 build 财务模型时会自动触发（SKILL.md 的 Quick Start 已指引）。Ship 前手动走一遍 `references/ib-deal-patterns.md` 底部的 checklist。

### PPT 模板

本次无改动。既有的 CJK 字宽引擎、content zone 边界、source footer 规范、QC pipeline 保留不变。

---

## [2026-04-20] — 初始版本

首次发布 PE/VC AI toolkit：

- **cathay-ppt-template**：国泰品牌 PPT 模板 + 12 版式 + 16 slide 模板 + CJK 文字引擎
- **cathay-excel-template**：13-sheet 财务模型 + 617 条预验证公式 + 10 项自动验证
- **skills/equity-research**：7-agent 辩论式深度研究
- **skills/market-sizing**：强制边界确认的自下而上 TAM sizing
- **skills/chain-screener**：产业链 mapping
- **skills/stock-screener**：5 层主题筛选
- **skills/stock-compare**：相对估值 / 同行对比

初始 commit：`2831e80`
