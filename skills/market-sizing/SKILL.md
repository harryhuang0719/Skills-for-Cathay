---
name: market-sizing
description: "Bottom-up market sizing, supply-demand balance analysis, and competitive landscape modeling. Use this skill whenever the user asks to: size a market or TAM/SAM/SOM, analyze supply-demand dynamics, map competitive landscape or market share, forecast industry pricing or volume, evaluate whether to invest in a sector/theme, do industry-level due diligence on capacity/utilization/ASP trends, or any variant of '帮我拆一下这个市场'. Also trigger when chain-screener or quant-analysis identifies an industry theme that needs deeper market-level validation. Output is a multi-sheet Excel workbook with full formula linkages showing the complete bottom-up model."
---

## ⛽ Pre-Flight Gate [R80][R93]

□ Step 0 — Phase 0 确认: 地理范围/产品范围/收入范围/时间基准/实体基准 — 用户必须确认后才执行
□ Step 1 — KB 检索: 行业相关知识库材料 [R31]
□ Step 2 — 数据源: 确认至少 1 个行业数据源可用 (web search / KB / FMP sector data)

⛔ 未完成 Phase 0 确认 → 中止

# Market Sizing + Supply-Demand + Competitive Landscape Skill

> **v3.2 — 2026-03-14 Validity-Enforced Engine Upgrade**
> V3.0 solved routing. V3.1 solved classification. V3.2 solves **execution integrity**.
> 核心认知：**Execution Integrity over Template Completeness — 模型是否有效，不由章节完整性决定，而由字段契约、公式落地、gate 触发三件事决定。**

---

## 核心理念

市场规模不是一个数字，是一组勾稽关系。但更重要的是：**每个输出必须能追溯到一条一致的因果链**。

表格连接 ≠ 经济因果。Demand/Supply/ASP/Competition 有公式引用，不代表它们之间有真正的驱动关系。

---

## ⚠️ Phase 0（强制）：建模前必须锁死的七件事

**在打开 Excel 之前，先回答这七个问题。没有答案不得开始建模。**

### 1. Market Boundary Box（五件套，必填）

| 字段          | 说明                                                         |
| ------------- | ------------------------------------------------------------ |
| Geography     | 全球 / 中国 / 北美 / 其他                                    |
| Product scope | 具体到哪个子品类（不是"光模块"而是"800G+数据中心光模块"）    |
| Revenue scope | hardware only / 含软件 / 含服务 / 全项目价值                 |
| Time basis    | shipment / sell-in / sell-through / installed base（选一个） |
| Entity basis  | brand / supplier / listed proxy（不同实体不能混用）          |

**任何输入数据都必须先过这个过滤器**——不符合边界定义的数据不能直接入主模型。

### 2. 价值链层级（只能选一层）

- 只能选：原材料 / 零部件 / 系统集成 / 终端产品 / 全项目价值
- **禁止跨层混算**：电芯厂和系统集成商不能在同一张竞争格局表里直接相加
- 若必须跨层，先建 layer bridge（每层占总价值的比例），再分别算

### 3. Driver Tree（因果顺序，建模主链）

```
1. Target base (population / devices / users)
2. Penetration / adoption rate
3. Units sold (realized shipments — 不是 theoretical capacity)
4. ASP → 必须二选一（见下）
5. Revenue = Units × ASP / unit_divisor
6. Supply constraint (capacity × utilization → cap check only)
7. Competitive allocation (demand-constrained, player-specific ASP)
8. Top-down reconciliation (gating function)
```

### 4. ASP 机制（必须二选一，不得混用）

- **Option A — Exogenous（analyst assumption）**：老实写 "ASP is exogenously assumed"，S/D 只做 consistency check，不装 driver
- **Option B — Endogenous（gap-driven）**：必须有公式 `ASP_t = max(PriceFloor, ASP_{t-1} × (1 + Elasticity × Gap%))`，弹性参数必须进入主计算链

### 5. 需求生成器（3年以上预测必须选一个）

任何超过3年的 forecast，必须有至少一个显式生成器，否则降级为 low-confidence：

| 生成器     | 公式                                                |
| ---------- | --------------------------------------------------- |
| 用户基数法 | target population × penetration × replacement cycle |
| 装机生态法 | installed base of compatible devices × attach rate  |
| 竞品替代法 | existing category users × substitution ratio        |
| 渠道法     | store reach × conversion × inventory turns          |

**禁止**：只写"品牌爆发/生态带动/教育成熟"而没有底层生成器的平滑增长曲线。

### 6. 竞争收入口径

- 正确：`Player Revenue = min(Player Supply, Player Allocated Demand) × Player ASP`
- **禁止**：`Player Revenue = Player Supply × Market Blended ASP`
- 硬验证：`sum(player revenue) <= total market revenue`，违反直接 FAIL

### 7. Top-down 验证（决策门，不是装饰）

必须输出：BU result / TD result / variance % / reason / pass or fail

| 偏差   | 处理                         |
| ------ | ---------------------------- |
| <10%   | ✅ 通过                      |
| 10-20% | ⚠️ 通过但必须解释            |
| >20%   | ❌ 必须修改或明确标注 caveat |

---

## 八条方法论铁律（v2.4，永久执行）

### Rule 1 — Unit Discipline（单位纪律）

- 每个模型开始前声明：Volume unit / Price unit / Revenue unit
- `/1000` 换算只允许在**一个地方**做，不允许多处分散
- 维度检查：GWh × $/kWh = $M（不是 $B）；K units × $/unit = $K（需 /1000 → $M）

### Rule 2 — Do Not Monetize Unsold Capacity

- player revenue 只能基于成交量，不能基于 nameplate / effective output
- 正确：`min(supply, allocated_demand) × player_ASP`
- 禁止：`supply × blended_ASP`

### Rule 3 — One Market, One Layer

- 同一张竞争格局表里的 player 必须在同一价值链层
- 若跨层，先建 layer bridge，标注每层占总价值比例

### Rule 4 — Narrative Must Map to Formula

- 说"X drives Y"，就必须有 `Y = f(X)` 的公式
- 没有公式就不要伪装成 driver-based model，标注"scenario assumption model"

### Rule 5 — Market ASP ≠ Player ASP

- blended ASP 只用于 total TAM
- player revenue 需要 player-specific mix 或 segment-specific price bridge

### Rule 6 — Top-Down Check Is a Gate

- TD check 必须输出 variance% + reason + pass/fail
- 不约束 forecast 的 TD check 不算 check

### Rule 7 — Evidence Grade Affects Model Weight

- Grade A：可直接作为核心 input
- Grade B：可作为锚点，需交叉验证
- Grade C：方向性参考，不能单点决定 forecast 斜率
- **禁止**：低质量假设和高质量数据拥有同等建模权重

### Rule 8 — Dashboard Polish ≤ Model Rigor

- Summary/Dashboard 必须显示主要输出是 assumption-driven 还是 data-driven
- 未解决的勾稽问题必须在 WARNINGS sheet 显示，不得隐藏在精美格式后面

---

## V3.2 执行完整性约束（Execution Integrity Rules，永久执行）

> 总原则：模型是否有效，不由章节完整性决定，而由 (1) required fields 完整且非空 (2) key formulas 真正落地 (3) gates 真正触发 三件事决定。

### Rule 9 — No Silent Fallback（禁止静默回退）

- 缺输入、缺公式、缺桥接时，不允许静默用默认值继续
- 只能：FAIL / SKIP with reason / DOWNGRADE with reason
- **禁止**：用 0 替代缺失 anchor、用手填替代未实现机制、用 CAGR 替代未建成 bridge

### Rule 10 — Formula Realization（公式落地）

- 核心输出（market revenue / ASP path / player revenue / adoption curve）必须由公式或可追溯链条生成
- 若为 manual override，必须在 `field_overrides` 中显式声明 type + reason
- 未声明的常数手填 → FAIL

### Rule 11 — Gate Realization（Gate 不可空跑）

- 每个 gate 必须有 trigger input + decision rule + fallback rule
- trigger input 缺失时，gate 不得默认 PASS，只能 FAIL / SKIP with reason

### Rule 12 — Dimension & Basis Integrity（量纲闭合）

- **Unit integrity**：volume × price = revenue 时单位必须可化简，K/M/B 缩放不得隐含跳跃
- **Basis integrity**：realized vs normalized / shipment vs sell-through 不得直接混算
- **Denominator integrity**：share / penetration / take-rate 的分子分母必须同层
- **Transformation integrity**：跨类型转换（用户→设备→收入）必须有 explicit bridge

### Rule 13 — Override Governance（Override 治理）

- 允许 manual override，但必须声明 scope / type / reason
- Override 分三类：`disclosed_anchor`（高权重）/ `expert_judgment`（降级）/ `temporary_placeholder`（不得进 final output）

### Rule 14 — Segment Heterogeneity（异质市场检测）

- 若 market boundary 内 billing_unit / price_formation / customer_type 显著不同，必须评估是否需要 segment-level archetype
- 使用 unified model 时必须提供 `segment_heterogeneity.justification`

### Rule 15 — Auditability by Construction（构造可审计）

- 模型必须输出 Audit sheet，包含：field completeness / formula realization status / gate trigger status / override log / issues by severity

### Field Taxonomy（字段契约分类）

| 类别               | 含义                                                                                                                          | 为空时       |
| ------------------ | ----------------------------------------------------------------------------------------------------------------------------- | ------------ |
| **A — Identity**   | 定义问题本身（sizing_objective / measurement_basis / realization_basis / time_horizon_type / billing_unit / market_boundary） | 模型不得开始 |
| **B — Mechanism**  | 定义计算方法（primary_archetype / formula_contract / generator_type / model_governance_bias）                                 | 模型不得发布 |
| **C — Commentary** | 辅助说明（archetype_rationale / notes）                                                                                       | 仅 warn      |

### Severity Taxonomy（严重性分级）

| 级别          | 含义                                         | 响应                        |
| ------------- | -------------------------------------------- | --------------------------- |
| L1 Cosmetic   | 仅影响可读性                                 | ignore                      |
| L2 Commentary | 说明不充分                                   | warn                        |
| L3 Mechanism  | 核心机制未实现                               | downgrade / not publishable |
| L4 Validity   | 身份不明 / 量纲错 / gate 空跑 / 关键结果伪装 | **FAIL**                    |

### Three-Layer Model Validity（三层有效性）

| 层         | 检查内容                                                            | 结果                |
| ---------- | ------------------------------------------------------------------- | ------------------- |
| Structural | identity fields / boundary / archetype / generator 存在性           | valid / invalid     |
| Mechanical | formula realization / gate trigger / unit integrity / bridge 真实性 | valid / invalid     |
| Economic   | archetype-formula 一致性 / price-volume 同层 / share 上限           | high / medium / low |

Structural / Mechanical 不过 → 不能发。Economic 可以有置信度差异。

---

## 分析框架（六步法）

### Step 1: 市场定义与边界界定

在开始任何测算前，必须先明确：

1. **产品/服务定义**：具体到什么品类？（例：不是"光模块"而是"800G+数据中心光模块"）
2. **地理边界**：全球/中国/北美？
3. **产业链位置**：上游原材料/中游零部件/下游终端？
4. **时间范围**：历史至少3年 + 预测3-5年
5. **计量单位**：收入（$）还是出货量（units）？通常两者都要

**输出**：一段清晰的market definition，作为Excel "Assumptions" sheet的第一行。

### Step 2: 需求侧底层拆解（Bottom-Up Demand）

核心公式：**市场规模 = Σ(子需求\_i × 单位用量\_i × ASP_i)**

拆解原则：

- **先拆子需求**：一个市场通常有2-5个核心下游应用场景
- **每个子需求继续拆**：终端保有量/新增量 × 单设备用量 × 替换周期
- **ASP单独建模**：ASP不是常数，它受供需缺口、技术迭代、规模效应影响
- **交叉验证**：Bottom-up加总 vs Top-down行业报告，偏差>15%需要解释

数据获取优先级：

1. 上市公司年报/investor day（最可靠）
2. 行业协会统计
3. 知识库已有研报（knowledge-retrieval）
4. Web search补充
5. 倒推估算（已知A公司market share 30%，收入$X → 行业规模 ≈ $X/0.3）

**关键**：每个数字都要标注来源。没有来源的数字标为 "E"（Estimate）。

### Step 3: 供给侧产能映射（Supply Mapping）

核心公式：**有效供给 = Σ(玩家\_j × 名义产能\_j × 开工率\_j)**

**重要**：Supply 只用于 consistency check 和 constraint，不直接驱动 revenue。

操作步骤：

1. 识别核心玩家（market share > 5%），其余归为"Others"
2. 产能数据来源：年报、investor day、行业报告
3. 区分名义产能 vs 有效产能
4. 产能扩张时间线（通常比公告晚6-18个月）

**关键约束**：各玩家产能加总 ≈ 行业总产能。如果不一致，需要解释差异。

### Step 4: 供需平衡与价格机制

**供需缺口 = 有效供给 - 实际需求**

分析维度：

1. 历史供需缺口 vs 历史ASP（建立价格对供需缺口的敏感度）
2. 库存周期（供需缺口不直接等于价格变化，中间有库存缓冲）
3. 价格弹性假设（短缺时弹性通常>1，过剩时<1）
4. 结构性 vs 周期性

**输出**：未来3-5年的ASP预测路径，基于供需缺口的量化推导（或明确标注为 analyst assumption）。

### Step 5: 竞争格局演变

**必须基于 realized revenue，不能基于 supply capacity。**

1. 市场集中度指标：CR3 / CR5 / HHI
2. 市占率变化趋势（至少3年历史）
3. 竞争壁垒评估：技术/规模/客户粘性/资本壁垒
4. 格局演变预测

**竞争收入分配逻辑**：

```
Player Shipment = Player Share × min(Total Demand, Total Supply)
Player Revenue = Player Shipment × Player-Specific ASP
```

### Step 6: 投资结论映射

1. 行业吸引力评分（基于供需结构）
2. 最佳投资窗口（基于供需缺口时间曲线）
3. 标的映射（链接到chain-screener）
4. 风险点

---

## Excel 输出规范

### 工作表结构（8 sheets）

| Sheet | 名称            | 内容                                                       |
| ----- | --------------- | ---------------------------------------------------------- |
| 0     | **WARNINGS**    | 自动检查结果，所有 FAIL/WARN 列表                          |
| 1     | **Summary**     | 仪表盘：关键指标、行业评分、投资结论、模型置信度           |
| 2     | **Assumptions** | 所有硬编码假设（蓝色字体）+ 数据来源 + Market Boundary Box |
| 3     | **Demand**      | 需求侧bottom-up拆解，子需求 × 用量 × ASP                   |
| 4     | **Supply**      | 各玩家产能、开工率、有效供给（consistency check only）     |
| 5     | **SD_Balance**  | 供需缺口计算、ASP路径（exogenous or endogenous，必须标注） |
| 6     | **Competition** | 市占率演变（基于realized revenue）、CR3/CR5、竞争壁垒      |
| 7     | **Data**        | 原始数据存档                                               |

### 格式规范

- **蓝色字体**：所有可调假设
- **黑色字体**：所有公式计算结果
- **绿色字体**：跨sheet引用
- **黄色背景**：关键假设
- **所有公式用Excel公式**，不在Python中计算后硬编码
- **Assumptions 中间计算行（segment total等）必须写硬值，不能写公式字符串**（openpyxl 陷阱：公式字符串不被其他公式计算）

---

## 执行工作流

### Phase 0: 边界确认（强制，不可跳过）

向用户确认 Market Boundary Box 五件套 + 价值链层级 + ASP 机制选择。

### Phase 1: 信息收集

```
1. knowledge-retrieval: 搜索知识库中的相关研报
2. web search: 补充关键数据点
3. 整理数据清单，标注每个数字的来源和可信度（Grade A/B/C）
```

### Phase 2: 模型构建

```
1. 先写 Driver Tree（因果顺序），再开表
2. 运行 scripts/generate_model.py
3. 填入收集到的数据和假设
4. Assumptions 中间计算行写硬值（非公式字符串）
```

### Phase 3: 质量检查（发送前强制，不通过不得发送）

**数学检查：**

- [ ] Revenue = Units × ASP / unit_divisor（维度正确）
- [ ] Demand 各子需求加总 = Total Demand（偏差 < 1%）
- [ ] Supply 各玩家加总 = Total Supply（精确相等）
- [ ] sum(player revenue) <= total market revenue（违反 = FAIL）
- [ ] Competition market share 加总 = 100%（±0.1%）
- [ ] CR3 ≤ CR5 ≤ 100%
- [ ] 无负数 volume / ASP（除非显式允许）

**逻辑检查：**

- [ ] ASP 机制已明确标注（exogenous or endogenous）
- [ ] 3年以上 forecast 有显式需求生成器
- [ ] Competition revenue 基于 realized demand，不基于 supply
- [ ] Player ASP 与 market blended ASP 有区分（或明确说明 mix 一致）
- [ ] 跨价值链层级的玩家已建 layer bridge 或已排除

**证据检查：**

- [ ] 每个数字有来源标注（不允许 "TBD"）
- [ ] Grade C 数据不单点决定核心 forecast 斜率
- [ ] Top-down check 输出 variance% + reason + pass/fail
- [ ] "Others" 产能 > 0

**口径检查：**

- [ ] 无 global capacity 混入 regional revenue
- [ ] 无 cumulative shipments 混入 annual sales
- [ ] 无 total company revenue 混入 hardware-only revenue
- [ ] 无 proxy ticker 冒充 player 主体

---

## 脚本说明

### scripts/generate_model.py (v2)

通过JSON配置文件生成Market Sizing Excel模型，支持**任意层级的需求拆解**和**非上市公司标注**。

#### 用法

```bash
python /path/to/market-sizing/scripts/generate_model.py \
  --config /path/to/config.json \
  --output /mnt/user-data/outputs/market_sizing_output.xlsx
```

#### JSON配置结构

```json
{
  "title": "Market Name",
  "years": ["2022", "2023", "2024", "2025E", "2026E"],
  "unit": "Tonnes",
  "revenue_unit": "$M",
  "asp_label": "Price ($/oz)",
  "unit_type": "stock",
  "annualization_factor": 1,
  "revenue_divisor": 1,
  "market_boundary": "明确写出：geography / product scope / revenue scope / time basis / entity basis",

  "demand": {
    "asp_mechanism": "exogenous",  // "exogenous" or "gap_driven"
    "demand_generator": "target_population × penetration",  // 需求生成器说明
    "segments": [...]
  },

  "supply": {
    "players": [...]
  }
}
```

#### 已知陷阱（openpyxl）

- Assumptions 中间计算行（如 segment volume total）必须写**硬值**，不能写公式字符串
- 原因：openpyxl 写入 `=B7+B8+B9` 后，其他公式读取时拿到字符串而非数值，导致加权 ASP = 0，Summary 全空
- 修复方式：在 Python 中预计算 total，写入数值

> 市场规模不是一个数字，是一组勾稽关系：需求侧的 **终端数量 × 单位用量 × 单价 = 市场规模**，供给侧的 **玩家产能 × 开工率 = 有效供给**，两者的缺口决定价格走向，价格走向反过来影响需求和供给的投资决策。

这个Skill的目标是把"XX市场有多大"这种模糊问题，拆解成一个**可验证、可预测、可投资**的量化模型。

---

## 分析框架（六步法）

### Step 1: 市场定义与边界界定

在开始任何测算前，必须先明确：

1. **产品/服务定义**：具体到什么品类？（例：不是"光模块"而是"800G+数据中心光模块"）
2. **地理边界**：全球/中国/北美？
3. **产业链位置**：上游原材料/中游零部件/下游终端？
4. **时间范围**：历史至少3年 + 预测3-5年
5. **计量单位**：收入（$）还是出货量（units）？通常两者都要

**输出**：一段清晰的market definition，作为Excel "Assumptions" sheet的第一行。

### Step 2: 需求侧底层拆解（Bottom-Up Demand）

核心公式：**市场规模 = Σ(子需求\_i × 单位用量\_i × ASP_i)**

拆解原则：

- **先拆子需求**：一个市场通常有2-5个核心下游应用场景（如：光模块 → 数据中心/电信/企业网）
- **每个子需求继续拆**：终端保有量/新增量 × 单设备用量 × 替换周期
- **ASP单独建模**：ASP不是常数，它受供需缺口、技术迭代、规模效应影响
- **交叉验证**：Bottom-up加总 vs Top-down行业报告，偏差>15%需要解释

数据获取优先级：

1. 上市公司年报/investor day（最可靠的出货量和收入拆分）
2. 行业协会统计（SEMI、WSTS、Dell'Oro等）
3. 知识库已有研报（knowledge-retrieval）
4. Web search补充（行业白皮书、咨询公司摘要）
5. 倒推估算（如：已知A公司market share 30%，收入$X → 行业规模 ≈ $X/0.3）

**关键**：每个数字都要标注来源。没有来源的数字标为 "E"（Estimate），并在Assumptions sheet说明估算逻辑。

### Step 3: 供给侧产能映射（Supply Mapping）

核心公式：**有效供给 = Σ(玩家\_j × 名义产能\_j × 开工率\_j)**

操作步骤：

1. **识别核心玩家**：列出market share > 5%的所有玩家，其余归为"Others"
2. **产能数据来源**：
   - 上市公司：年报、investor day、earnings call（搜索"capacity", "capex", "expansion"）
   - 非上市公司：行业报告、新闻稿、政府审批公告
3. **区分名义产能 vs 有效产能**：良率、维护停机、产品mix都会影响有效产出
4. **产能扩张时间线**：在建产能的投产时间（通常比公告晚6-18个月）
5. **新进入者**：识别已宣布进入但尚未形成产能的玩家

**关键约束**：各玩家产能加总 ≈ 行业总产能，各玩家revenue加总 ≈ 行业总收入。如果不一致，需要解释差异（通常是"Others"的估算问题）。

### Step 4: 供需平衡与价格机制

这是整个模型最有投资价值的部分。

**供需缺口 = 有效供给 - 实际需求**

分析维度：

1. **历史供需缺口 vs 历史ASP**：建立价格对供需缺口的敏感度
   - 画出历史散点图：X轴=供需缺口率(%)，Y轴=ASP同比变化(%)
   - 识别是否存在非线性关系（如：缺口>10%时价格暴涨）
2. **库存周期**：供需缺口不直接等于价格变化，中间有库存缓冲
   - 渠道库存天数（如有数据）
   - 库存同比变化方向
3. **价格弹性假设**：
   - 供给短缺时：价格上涨弹性（通常>1，因为恐慌性采购）
   - 供给过剩时：价格下跌弹性（通常<1，因为寡头定价纪律）
4. **结构性 vs 周期性**：
   - 结构性短缺（如：技术瓶颈导致产能无法快速扩张）→ ASP持续上行
   - 周期性短缺（如：需求脉冲 + 产能投资周期）→ ASP先涨后跌

**输出**：未来3-5年的ASP预测路径，基于供需缺口的量化推导，不是拍脑袋。

### Step 5: 竞争格局演变

在Step 3的基础上进一步分析：

1. **市场集中度指标**：
   - CR3 / CR5（前3/5名市占率之和）
   - HHI指数（可选，寡头市场更有意义）
2. **市占率变化趋势**：至少3年的历史market share演变
3. **竞争壁垒评估**：
   - 技术壁垒（良率、专利、know-how）
   - 规模壁垒（成本曲线陡峭程度）
   - 客户粘性（认证周期、switching cost）
   - 资本壁垒（产能投资额 vs 回报周期）
4. **格局演变预测**：
   - 哪些玩家在扩产？扩产速度 vs 行业增速
   - 新进入者的实际威胁（有产能规划 vs 只有PPT）
   - 是否有整合（M&A）趋势

### Step 6: 投资结论映射

将以上分析收敛为投资判断：

1. **行业吸引力评分**（基于供需结构）：
   - 供不应求 + 集中度高 + 扩产慢 = ⭐⭐⭐⭐⭐
   - 供需平衡 + 集中度中 = ⭐⭐⭐
   - 供过于求 + 新进入者多 = ⭐
2. **最佳投资窗口**：基于供需缺口的时间曲线
3. **标的映射**：哪些上市公司是这个行业的最佳代理？（链接到chain-screener）
4. **风险点**：供需反转时间点、技术替代风险、政策风险

---

## Excel 输出规范

### 工作表结构（6 sheets）

| Sheet | 名称            | 内容                                     |
| ----- | --------------- | ---------------------------------------- |
| 1     | **Summary**     | 仪表盘：关键指标、行业评分、投资结论     |
| 2     | **Assumptions** | 所有硬编码假设（蓝色字体）+ 数据来源     |
| 3     | **Demand**      | 需求侧bottom-up拆解，子需求 × 用量 × ASP |
| 4     | **Supply**      | 各玩家产能、开工率、有效供给、扩产计划   |
| 5     | **SD_Balance**  | 供需缺口计算、历史价格关联、ASP预测      |
| 6     | **Competition** | 市占率演变、CR3/CR5、竞争壁垒评估        |

### 格式规范

遵循 `/mnt/skills/public/xlsx/SKILL.md` 的全部格式要求，额外强调：

- **蓝色字体**：所有可调假设（growth rate, ASP assumption, utilization rate等）
- **黑色字体**：所有公式计算结果
- **绿色字体**：跨sheet引用（如Supply sheet引用Assumptions sheet的开工率假设）
- **黄色背景**：关键假设（需要用户特别关注的）
- **所有公式用Excel公式**，不在Python中计算后硬编码
- **年份格式为文本**，不要出现 "2,024"
- **勾稽关系检查行**：在关键位置加入 "Check: Bottom-up Total vs Top-down" 行，公式为差异百分比

### 命名规范

文件名：`market_sizing_{行业关键词}_{日期}.xlsx`  
示例：`market_sizing_800G_optical_20260313.xlsx`

---

## 执行工作流

当触发此Skill时，按以下顺序执行：

### Phase 0: 边界确认（强制，不可跳过）⚠️

**在做任何信息收集或建模之前，必须先向用户确认以下边界条件：**

```
1. 地理范围：全球 / 中国 / 北美 / 亚太 / 其他？
2. 产业链位置：上游原材料 / 中游零部件 / 下游终端 / 全链条？
3. 产品口径：具体到哪个子品类？（例：不是"光模块"而是"800G+数据中心光模块"）
4. 时间范围：历史起点年份 + 预测终点年份？
5. 计量单位偏好：收入（$）/ 出货量（units）/ 产能（MW/GW等）？
```

**格式：** 用一条简短消息列出以上5点，等用户确认后再进入Phase 1。
**禁止：** 不得假设任何边界条件后直接开始建模。

### Phase 1: 信息收集（Information Gathering）

```
1. 明确市场定义（与用户确认边界）
2. knowledge-retrieval: 搜索知识库中的相关研报
3. web search: 补充关键数据点
   - "{行业} market size 2024 2025"
   - "{行业} capacity expansion {top players}"
   - "{top player} investor day capacity utilization"
   - "{行业} supply demand balance"
4. 如果用户提供了额外数据/研报，优先使用
5. 整理数据清单，标注每个数字的来源和可信度
```

### Phase 2: 模型构建（Model Building）

```
1. 读取 /mnt/skills/public/xlsx/SKILL.md 的格式规范
2. 运行 scripts/generate_model.py（本Skill自带）
3. 填入收集到的数据和假设
4. 运行 /mnt/skills/public/xlsx/scripts/recalc.py 重算公式
5. 检查公式错误，修复后再次recalc
```

### Phase 3: 质量检查（QA Checklist）

**⚠️ v2.1 强化：完整检查清单见 `memory/shared/skill_market_sizing.md` Phase 4.5 部分。**
**核心原则：一个错误的数字比一个空白的单元格更有害。**

在交付前必须验证：

- [ ] 需求侧各子需求加总 = 总需求（Demand sheet内部勾稽）
- [ ] 各玩家产能加总 ≈ 行业总产能（Supply sheet内部勾稽，subset 玩家已排除）
- [ ] 各玩家revenue加总 ≈ 行业总收入（Supply × ASP ≈ Market Size）
- [ ] 历史数据与公开来源可交叉验证
- [ ] ASP预测路径有供需缺口的量化支撑
- [ ] 竞争格局百分比加总 = 100%
- [ ] 所有假设标注来源（不允许 "TBD"）
- [ ] 无Excel公式错误
- [ ] "Others" 产能 > 0（v2.1 硬红线）
- [ ] 没有数量级错误（每个数字必须 web search 验证，v2.1 规则 3A）
- [ ] Summary sheet 全部为公式引用，无手填硬编码（v2.1 硬红线）
- [ ] 大宗商品：Price Mechanism 参数已填写（v2.1 大宗商品规则）
- [ ] 大宗商品：subset 玩家已标注 is_subset_of，无重复计算（v2.1）

---

## 脚本说明

### scripts/generate_model.py (v2)

通过JSON配置文件生成Market Sizing Excel模型，支持**任意层级的需求拆解**和**非上市公司标注**。

#### 用法

```bash
python /path/to/market-sizing/scripts/generate_model.py \
  --config /path/to/config.json \
  --output /mnt/user-data/outputs/market_sizing_output.xlsx
```

#### JSON配置结构

```json
{
  "title": "Market Name",
  "years": ["2022", "2023", "2024", "2025E", "2026E"],
  "unit": "Tonnes", // 需求量单位
  "revenue_unit": "$M", // 收入单位
  "asp_label": "Price ($/oz)", // ASP展示名
  "unit_type": "stock", // "stock" (年度量) 或 "flow" (日度流量如 mb/d)
  "annualization_factor": 1, // flow 型设 365，stock 型设 1
  "revenue_divisor": 1, // 如需单位转换（如 $M → $B）设 1000

  "demand": {
    "segments": [
      {
        "name": "Central Bank Purchases",
        "sub_segments": [
          // ← 可以有任意多个子需求
          { "name": "PBOC (China)", "source": "WGC report" },
          { "name": "RBI (India)" },
          { "name": "Other CBs" }
        ]
      },
      {
        "name": "Technology/Industrial" // ← 没有sub_segments则作为单层处理
      }
    ]
  },

  "supply": {
    "players": [
      {
        "name": "Newmont",
        "listed": true, // ← 上市公司
        "ticker": "NEM",
        "source_quality": "A", // A=年报验证 B=IR报告 C=行业估算 D=粗略估计
        "notes": "World's largest gold miner"
      },
      {
        "name": "Tsingshan (青山控股)",
        "listed": false, // ← 非上市公司
        "ticker": null,
        "source_quality": "C", // 数据可信度较低
        "notes": "PRIVATE. Capacity from Wood Mackenzie + Indonesian govt data.",
        "is_subset_of": null // ← 如果是某国家汇总的子集，填父级名称
      }
    ]
  }
}
```

#### 关键特性

1. **多层需求拆解**：segments → sub_segments，每个sub_segment在Assumptions中有独立的volume输入行，Demand sheet自动加总
2. **非上市公司标注**：
   - 绿色背景 = 上市公司（数据可信）
   - 橙色背景 = 非上市/私有公司（数据需谨慎）
   - Source Quality等级：A（年报验证）> B（IR报告）> C（行业估算）> D（粗略估计）
   - notes字段说明数据来源和局限性
3. **Player Profile表**：Supply sheet顶部自动生成玩家概况表（状态/ticker/来源等级/备注）
4. **New Entrant Tracker**：Competition sheet底部预留新进入者/扩产追踪表
5. **完整公式链**：Assumptions → Demand/Supply → SD_Balance → Competition → Summary，全部用Excel公式

#### 配置示例

参见：

- `examples/config_idc.json` — 中国IDC托管市场（v8验证通过，0 FAIL，2026-03-14）
  - 5个需求段 × 2个子段，7个供应商（全部上市），利用率走势各异（独立/电信/AI各不同）
  - 可作为任何服务类市场（按MW/Rack/sqm计量）的参考模板

如果不使用脚本（例如数据结构不适合模板），可以直接用openpyxl手写，但必须遵循上述6-sheet结构和格式规范。

### 快速运行（任意行业）

```bash
cd ~/.openclaw/workspace
python3 skills/market-sizing/scripts/run_market_sizing.py \
  --industry "800G数据中心光模块" \
  --geography "全球" \
  --scope "中游零部件" \
  --time-horizon 5 \
  --currency USD \
  --output skills/market-sizing/output/market_sizing_800G_optical_$(date +%Y%m%d).xlsx
```

**常用行业参数示例：**

| 行业       | `--industry`                      | `--geography` | `--scope` |
| ---------- | --------------------------------- | ------------- | --------- |
| 中国IDC    | 中国IDC（互联网数据中心）托管服务 | 中国          | 中游      |
| 黄金       | 黄金                              | 全球          | 全链条    |
| HBM内存    | HBM高带宽内存                     | 全球          | 中游      |
| 人形机器人 | 人形机器人                        | 全球          | 下游终端  |
| 光伏组件   | 光伏组件                          | 全球          | 中游      |
