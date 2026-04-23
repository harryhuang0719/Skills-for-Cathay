# 8 阶段方法论 — IB/PE Deck 生产流水线

> 改编自麦肯锡/德勤咨询流程（aham-ppt），适配 IB/PE 场景下的 IC memo / pitch deck / 投资人路演稿。
>
> **核心纪律：先想清楚再动手。** Ghost Deck 没有落纸前不要跳到版式设计。
>
> 什么时候走这套流程：篇幅 ≥ 8 页、或交付对象是合伙人 / LP / 投委会。
> 内部 quick take 3-5 页 deck 可以跳过阶段 3-6。

---

## 阶段 1 · 规范加载

**目的**：确认品牌/版式基线，不要自己造。

- 读 `SKILL.md` 的 Brand Identity / Global Style Constants
- 读 `references/design-philosophy.md` 的 4 支柱 + 禁用清单
- 读 `references/action-title-rules.md` 确认每页标题规则

**输出**：确认用 Cathay 主色 (Maroon `#800000` + Gold `#E8B012`)、楷体+Calibri 字体栈、可见网格。

---

## 阶段 2 · 材料解析

**目的**：把原始素材里所有可能进 deck 的信息提干净，**不做筛选**。

常见原始素材：
- CIM / teaser / 公司官网 / 访谈纪要
- 财务模型（来自 cathay-excel-template，取 Key Assumptions / Returns）
- 尽调材料、管理层访谈、行业研究
- 二级市场 comps（来自 equity-research / stock-compare）

**输出**：一个 markdown 大纲，含所有候选事实 + 每条来源标注（页码/文件名）。来源缺失的事实**不入后续阶段**。

---

## 阶段 3 · 论点提炼（投资逻辑金字塔）

**目的**：从大量素材里收敛到 1 个主张 + 3-5 个支柱 + 每个支柱 2-3 条证据。

典型 IC memo 金字塔：

```
主张：Sealien 是海洋机器人细分赛道的领导者，5 年 IRR 35%+

支柱 1：市场 TAM 800 亿 + CAGR 25%，头部效应强
  - 证据 A：全球装机量 2023-2028 CAGR 21%（数据来源）
  - 证据 B：海上油气 / 风电运维两个场景叠加
  - 证据 C：国产替代缺口 200 亿

支柱 2：技术壁垒（具身智能 + 水下定位双栈）
  - 证据 A-C...

支柱 3：商业化路径清晰（B端大客户锁定 + 服务化转型）
  - 证据 A-C...

支柱 4：估值 30 亿合理 + exit case 2030 年 50x PE 路径
  - 证据 A-C...
```

**禁忌**：
- 支柱 > 5 条 — LP 记不住，且说明论点没有被真正收敛
- 证据数字跟模型对不上 — 先跑 model，再回来改 deck
- 支柱互相重叠 — 每个支柱应该是独立 dimension（市场 / 技术 / 商业化 / 估值 / 团队）

---

## 阶段 4 · Ghost Deck（叙事骨架）

**目的**：**不画任何版式，只写每页 Action Title。** 这是整个流程最重要的一步。

Ghost Deck 范例：

```
p1  [封面] Project Sealien — 海洋机器人投资推介
p2  [目录]
p3  执行摘要：Cathay 拟投 2 亿，对应 Round A 30 亿 pre-money，目标 2030 年 50x 退出
p4  海洋机器人市场：TAM 800 亿、CAGR 25%、国产替代空间 200 亿
p5  Sealien 切入点：海上风电运维 + 深海资源探测双场景
p6  公司快照：2021 成立、CEO 背景、团队 80 人、研发占 60%
p7  产品矩阵：水下机器人 3 款 + 控制软件 + 具身大模型
p8  商业模式：硬件 + 服务化收费双线，2025 硬件 65% 服务 35%
p9  财务表现：2023-2025 收入 CAGR 180%，毛利提升至 42%
p10 预测：2026-2030 收入 CAGR 45%，2028 扭亏
p11 投资亮点 A：技术壁垒（具身 + 水下 SLAM）
p12 投资亮点 B：客户锁定（中海油 / 国电 / 三峡）
p13 投资亮点 C：商业化拐点（2026 服务化收入占比突破 40%）
p14 风险：技术迭代、订单集中度、毛利下行（已对冲项列明）
p15 交易结构：Cathay 2 亿 / Round A 总 4 亿 / 30 亿 pre-money
p16 退出路径：2030 IPO Base 25x / Bull 50x
p17 IRR/MOIC：Base 28% / 4.2x · Bull 41% / 6.8x
p18 Next Steps：尽调时间表 + 关键 milestones
```

**检查**：
- 每页标题是**完整结论句**（详见 `action-title-rules.md`）
- Ghost Deck 读完一遍，LP 能复述主要论点就合格
- 跨页数字是否自洽（TAM 800 亿 vs 市占率 % vs 公司收入）

**输出**：一份纯 markdown 的 Ghost Deck。此时**绝对不要开始画版式**。

---

## 阶段 5 · 大纲与版式规划

**目的**：给每页 Ghost Title 匹配版式类型（选用 `slide_templates.py` 里 T1-T16）。

常见 IB/PE 场景版式对照：

| 内容 | 推荐版式 |
|------|---------|
| 封面 | T1（Red Title Cover） |
| 目录 | T2（Red Band TOC）|
| 市场/TAM 分析 | T8（Data Card + 图表）或 T11（Donut）|
| 公司快照 | T3（双栏内容）|
| 产品矩阵 | T6（3 卡片）|
| 商业模式 | T10（Flow Diagram 或 T13 Funnel）|
| 财务预测 | smart_table + T8 数据卡 |
| 投资亮点 | T4（三栏对比）|
| 风险 | T12（Before/After）或 T14（SWOT）|
| 交易结构 | smart_table（primary） |
| Returns | smart_table + T8 |
| 退出路径 | T15（Waterfall） |

**输出**：表格，每行 = 一页，列 = [页码 | Ghost Title | 版式 T# | 数据依赖] 。

---

## 阶段 6 · 样稿确认

**目的**：选 3-5 页代表性页做样稿，和用户对齐视觉方向**再**批量生产。

推荐样稿页：
- 封面
- 1 页"高密度表格"（财务预测 / Comps）
- 1 页"视觉叙事"（商业模式图 / 产业链）
- 1 页"投资亮点"三栏

样稿确认点：
- 字号层级是否够大（banker 投屏看得清）
- 数字对齐和格式（千分符 / 百分号 / 负数括号）
- 颜色是否只用 Maroon 一个强调色（不要第二装饰彩）

**如果用户要求动主色/字体/版式整体换** → 回到阶段 1，不要在样稿层打补丁。

---

## 阶段 7 · 逐页设计输出

**目的**：批量生产剩余页面。

标准流程：
1. 全部用 `smart_textbox()` / `smart_table()` — 永远不要直接创建 textbox
2. 表格列宽严格对齐（数字列右对齐 + 等宽字体）
3. 每页 Source footer 7pt @ y=182mm
4. 视觉页（封面 / Divider / 商业模式图 / chart）**可以**走 `svg_embed.py` 通道：出 SVG → 渲染 PNG → `slide.shapes.add_picture(...)`。**但这些页会失去可编辑性**，只对"成品图"类页使用。
5. 财务表格页、数字密集页**必须**用 python-pptx 原生 Table，不要走 SVG 嵌入

**禁忌**：
- 为了速度跳过 `smart_textbox` 直接用 `add_textbox` — CJK 字宽会崩
- 把 Comps 表走 SVG 通道（不可编辑，banker 无法改列/改数字）

---

## 阶段 8 · 质检交付

**目的**：在发出前把能自动检测的错误全修掉。

标准 QC pipeline（已在 `lib/qc_automation.py`）：
1. 8 条 guard rails（字号、密度、bullet、overlap...）
2. 4-stage autofix
3. 文字溢出全量扫描
4. PDF 导出人工 review

**额外 IB/PE 特有检查**（建议人工过一遍）：
- [ ] 跨页数字自洽（TAM / 市占率 / 收入 / Returns 是否对齐模型）
- [ ] Action Title 都是完整结论句（详见 `action-title-rules.md`）
- [ ] Source footer 每页都有（至少 company / 研报 / 管理层访谈 三选一）
- [ ] 禁用词扫描（赋能 / 颠覆 / 一站式 / 显著 / 大幅 — 详见 `design-philosophy.md`）
- [ ] 所有具体数字有明确来源（scenario 标记 Base/Bull）

**交付**：`.pptx` + `.pdf`（PDF 用 soffice 转出），git 版本化保存。

---

## 流程短路：Quick Take 模式（3-5 页，内部用）

如果是内部 team 快速讨论的 3-5 页 quick take：

跳过阶段 3、5、6。直接从阶段 2 → 阶段 4（Ghost Deck） → 阶段 7。

阶段 1、8 永远不要跳过。
