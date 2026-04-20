# Industry Valuation Paradigms

## Routing Logic

Priority: exact industry match → sector match → generic fallback.

Given a company's `sector` and `industry` (from FMP profile or web search):
1. Check if `industry` substring-matches any paradigm's industries list
2. If not, check if `sector` matches any paradigm's sectors list
3. If not, use `generic` paradigm

**Quick mode**: Use the top-2 highest-weighted valuation methods.
**Deep mode**: Use all methods in the paradigm.

---

## 1. Mining Paradigm (矿业范式)

**Sectors**: Basic Materials
**Industries**: Gold, Silver, Copper, Uranium, Lithium, Other Precious Metals, Other Industrial Metals, Coking Coal, Thermal Coal, Steel

**Key Metrics**: NAV/储量寿命, AISC vs 商品价格, 储量替换率, EV/Resource, 产量增长

**Valuation Methods**:
| Method | Weight |
|--------|--------|
| NAV (Net Asset Value) | 40% |
| EV/EBITDA (周期调整) | 30% |
| DCF (储量寿命) | 30% |

**Industry Expert Prompt**:
分析时必须关注:
1. 储量寿命和品位趋势（是否在衰竭？）
2. AISC（全维持成本）vs 当前商品价格 — 这决定了利润空间
3. 长协 vs 现货比例 — 收入可见性
4. 资本开支周期位置 — 是在扩产还是维持？
5. 地缘政治风险（矿区所在国）
不要用通用 P/E 估值矿业公司，必须用 NAV 或 EV/Resource。

**Risk Factors**: 商品价格暴跌, 矿区政治风险/资源国有化, 储量品位下降, 环保法规收紧, 资本开支超支

---

## 2. Tech Growth Paradigm (科技成长范式)

**Sectors**: Technology
**Industries**: Software - Application, Software - Infrastructure, Semiconductors, Information Technology Services, Electronic Components

**Key Metrics**: TAM/渗透率, Rule of 40, NRR/ARR增速, FCF Margin, R&D/Revenue

**Valuation Methods**:
| Method | Weight |
|--------|--------|
| DCF (高增长) | 35% |
| EV/Revenue (成长倍数) | 30% |
| P/E (盈利公司) 或 EV/GP | 35% |

**Industry Expert Prompt**:
分析时必须关注:
1. TAM（总可寻址市场）和当前渗透率 — 增长天花板在哪？
2. Rule of 40（收入增速 + FCF margin）— 衡量增长质量
3. NRR（净收入留存率）— 客户粘性和扩展能力
4. 竞争壁垒（网络效应/转换成本/数据护城河）
5. AI 对业务的影响（正面还是颠覆？）
对于亏损科技公司，用 EV/Revenue 或 EV/GP，不要用 P/E。

**Risk Factors**: TAM 见顶/渗透率饱和, 竞争加剧导致定价压力, AI 颠覆现有商业模式, 客户集中度风险, 监管风险

---

## 3. Shipping Paradigm (航运范式)

**Sectors**: Industrials, Energy
**Industries**: Marine Shipping, Oil & Gas Midstream, Shipping

**Key Metrics**: P/NAV, 运价周期位置, 船龄/船队结构, 日租金 vs 盈亏平衡, 新船订单/船队比

**Valuation Methods**:
| Method | Weight |
|--------|--------|
| P/NAV (船队净资产) | 45% |
| 正常化盈利 P/E | 30% |
| Dividend Yield (周期高点) | 25% |

**Industry Expert Prompt**:
分析时必须关注:
1. P/NAV — 当前市值 vs 船队净资产价值（二手船价）
2. 运价周期位置 — 当前运价在历史什么分位？
3. 供给侧：新船订单/现有船队比 — 未来供给压力
4. 需求侧：贸易量增长、航线结构变化
5. 船龄结构 — 老船淘汰带来的供给收缩
航运是强周期行业，不要在周期高点用当前盈利做 P/E 估值。

**Risk Factors**: 运价周期下行, 新船交付潮, 地缘政治改变航线, 环保法规（IMO 2030）, 油价波动影响成本

---

## 4. Oil & Gas Paradigm (油气范式)

**Sectors**: Energy
**Industries**: Oil & Gas Integrated, Oil & Gas E&P, Oil & Gas Refining, Oil & Gas Equipment & Services

**Key Metrics**: 储量寿命(R/P), Finding Cost, 盈亏平衡油价, FCF Yield, 资本纪律

**Valuation Methods**:
| Method | Weight |
|--------|--------|
| NAV (储量价值) | 35% |
| EV/EBITDA (正常化) | 35% |
| FCF Yield | 30% |

**Industry Expert Prompt**:
分析时必须关注:
1. 盈亏平衡油价 — 在什么油价下公司 FCF 为正？
2. 储量替换率和 Finding Cost — 可持续性
3. 资本纪律 — 是否在高油价时过度扩张？
4. 能源转型风险 — 长期需求前景
5. 股东回报（回购+分红）vs 再投资比例

**Risk Factors**: 油价暴跌, 能源转型加速, 地缘政治供给中断, 资本开支失控, ESG 压力导致融资困难

---

## 5. Cyclical Consumer Paradigm (周期消费范式)

**Sectors**: Consumer Cyclical
**Industries**: (catch-all for sector — any Consumer Cyclical industry)

**Key Metrics**: 正常化盈利, 周期位置评分, 库存周期, 消费者信心, 同店销售增速

**Valuation Methods**:
| Method | Weight |
|--------|--------|
| P/E (正常化盈利) | 40% |
| EV/EBITDA | 30% |
| DCF | 30% |

**Industry Expert Prompt**:
分析时必须关注:
1. 正常化盈利 — 不要用周期高点/低点的盈利做估值
2. 库存周期位置 — 去库存还是补库存？
3. 消费者支出趋势和信心指数
4. 品牌力和定价权
5. 渠道库存健康度

**Risk Factors**: 经济衰退导致需求下滑, 库存积压, 消费降级, 原材料成本上涨, 竞争加剧

---

## 6. Financials Paradigm (金融范式)

**Sectors**: Financial Services
**Industries**: Banks - Regional, Banks - Diversified, Insurance - Life, Insurance - Property & Casualty, Insurance - Diversified, Insurance - Specialty, Insurance - Reinsurance, Asset Management, Capital Markets, Credit Services, Financial Data & Stock Exchanges, Mortgage Finance, Financial Conglomerates

**Key Metrics**: P/TBV, ROE, ROA, NIM, Cost/Income Ratio, CET1 Ratio, NPL Ratio, Book Value Growth, Dividend Yield

**Valuation Methods**:
| Method | Weight |
|--------|--------|
| P/TBV (相对 ROE) | 35% |
| Dividend Discount Model (DDM) | 30% |
| P/E (正常化) | 20% |
| Residual Income Model | 15% |

**Industry Expert Prompt**:
分析时必须关注:
1. ROE vs COE (cost of equity) — ROE > COE 才能创造价值，P/TBV 应 > 1.0x
2. NIM（净息差）趋势 — 利率周期对银行盈利的影响是核心
3. 资产质量 — NPL ratio趋势、拨备覆盖率、信贷成本
4. 资本充足率 — CET1 ratio、RWA增速、监管要求缓冲
5. Cost/Income ratio — 运营效率和规模经济
6. 保险公司关注 Combined Ratio (< 100% = 承保盈利)、Investment Yield、Reserve Development
7. 资管公司关注 AUM增速、fee rate趋势、Performance fees占比
金融公司不能用 EV/EBITDA（资本结构特殊），必须用 P/TBV 或 P/E。银行的"收入"包含利息收入，不适合 EV/Revenue。

**Risk Factors**: 利率方向逆转, 信贷周期恶化/NPL飙升, 监管收紧（资本要求/压力测试）, 系统性金融风险传染, 金融科技颠覆

---

## 7. Healthcare & Biotech Paradigm (医疗健康范式)

**Sectors**: Healthcare
**Industries**: Biotechnology, Drug Manufacturers - General, Drug Manufacturers - Specialty & Generic, Medical Devices, Medical Instruments & Supplies, Diagnostics & Research, Health Information Services, Healthcare Plans, Medical Care Facilities, Pharmaceutical Retailers

**Key Metrics**: Pipeline rNPV, Patent Cliff Exposure, R&D Productivity, Revenue/Drug, LOE (Loss of Exclusivity) Timeline

**Valuation Methods**:
| Method | Weight | Applicable To |
|--------|--------|---------------|
| rNPV (risk-adjusted NPV of pipeline) | 40% | Biotech, Drug Manufacturers |
| DCF (盈利公司) | 30% | Profitable pharma, devices, plans |
| EV/EBITDA | 20% | Profitable companies |
| Sum-of-the-Parts (existing drugs + pipeline) | 10% | Diversified pharma |

**Industry Expert Prompt**:
分析时必须关注:
1. **Pipeline价值** — 用rNPV给每个在研药物赋概率加权估值。Phase I (~10%), Phase II (~25%), Phase III (~50%), NDA/BLA (~80%), Approved (~95%)
2. **Patent Cliff** — 核心产品的专利到期时间表、仿制药/生物类似药威胁。LOE前3年收入通常下降50-80%
3. **R&D Productivity** — 研发投入vs NME/BLA获批数量、临床成功率 vs 行业平均
4. **Reimbursement & Pricing** — 政府定价政策风险（IRA、Medicare negotiation）、PBM渠道谈判
5. **M&A作为增长引擎** — 大型pharma经常通过收购补充pipeline，关注BD历史和整合能力
6. **Medical Devices** — 用DCF + EV/EBITDA，关注procedure volume增速、ASP趋势、regulatory pathway (510k vs PMA)
7. **Healthcare Plans (保险)** — 用P/E + Medical Loss Ratio (MLR)，关注membership增速、Star Ratings、政策风险
Pre-revenue biotech不能用P/E或EV/EBITDA，必须用rNPV。有盈利的pharma/devices可以用DCF。

**Risk Factors**: 临床试验失败, 专利悬崖/仿制药竞争, 药品定价政策收紧, FDA审批延迟/拒绝, 并购整合风险

---

## 8. Consumer Staples & Utilities Paradigm (防御消费与公用事业范式)

**Sectors**: Consumer Defensive, Utilities
**Industries**: Household & Personal Products, Packaged Foods, Beverages - Non-Alcoholic, Beverages - Alcoholic, Tobacco, Food Distribution, Farm Products, Utilities - Regulated Electric, Utilities - Regulated Gas, Utilities - Diversified, Utilities - Renewable, Utilities - Independent Power Producers

**Key Metrics**: Organic Revenue Growth, Pricing Power (价格弹性), Dividend Yield, Payout Ratio, FCF Stability, Regulated ROE (公用事业)

**Valuation Methods**:
| Method | Weight |
|--------|--------|
| DCF (稳定增长) | 35% |
| P/E (相对历史) | 30% |
| Dividend Yield / DDM | 20% |
| EV/EBITDA | 15% |

**Industry Expert Prompt**:
分析时必须关注:
1. **Organic Growth** — 拆分volume和pricing。消费必需品靠定价权驱动增长，volume增速通常 < 2%
2. **品牌溢价和定价权** — 消费者对价格变化的敏感度（弹性），私标(private label)份额趋势
3. **输入成本** — 原材料（农产品、包装、能源）成本占比和对冲策略
4. **股东回报** — 这些是经典的"dividend play"，关注payout ratio可持续性、股息增长历史、回购
5. **Utility特殊分析**:
   - Rate base增速 — 监管审批的CapEx计划是增长引擎
   - Allowed ROE vs Earned ROE — 是否在赚取监管允许的回报？
   - Regulatory environment — 所在州的监管友好度（constructive vs hostile）
   - Fuel mix和clean energy转型成本
6. **估值锚** — 消费必需品通常交易在20-25x P/E（稳定溢价），Utilities交易在15-18x P/E。大幅偏离需要解释
不要期望高增长（GDP+1-2%是常态）。估值核心是稳定性和分红能力，不是成长性。

**Risk Factors**: 消费降级/私标替代, 原材料成本飙升无法转嫁, 利率上升压低防御板块估值, 监管费率审批不及预期(utilities), 新兴品牌/DTC渠道颠覆

---

## 9. Generic Paradigm (通用范式)

**Sectors**: (fallback for any unmatched sector)
**Industries**: (fallback)

**Key Metrics**: P/E, P/S, P/B, EV/EBITDA, FCF Yield

**Valuation Methods**:
| Method | Weight |
|--------|--------|
| P/E 相对估值 | 35% |
| DCF | 35% |
| EV/EBITDA | 30% |

**Industry Expert Prompt**:
使用多重估值方法交叉验证，关注:
1. 盈利质量和可持续性
2. 竞争优势和护城河
3. 管理层资本配置能力
4. 行业增长前景

**Risk Factors**: 宏观经济下行, 竞争加剧, 管理层风险, 估值过高
