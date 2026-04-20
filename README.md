# 🦅 Skills for Cathay

这是我整理的一套 AI Skills，专门给咱们做一级市场 PE/VC 工作用的。涵盖了做 deck、建模型、做 market sizing、产业链分析这些日常高频场景。

简单来说：把这些 skill 喂给任何 AI agent，就能直接用自然语言让它帮你出 PPT、建财务模型、做行业分析。省掉大量重复劳动，把时间花在判断上而不是格式上。

---

## 🌍 兼容哪些平台

这些 skill 本质上是 **结构化的 Markdown 指令 + Python 辅助库**，不绑定任何特定平台。以下环境都能直接用：

| 平台 | 怎么用 |
|------|--------|
| 🟣 **Claude Code** | 放到 `~/.claude/skills/` 目录，自动识别 |
| 🐾 **OpenClaw Agent** | 放到 `~/.openclaw/workspace/skills/` 目录 |
| 🟢 **OpenAI Codex CLI** | 作为 instructions 文件引用（`AGENTS.md` 或 `--instructions`） |
| 🔵 **Cursor / Windsurf** | 放到项目 `.cursor/rules/` 或作为 context 引入 |
| 🔶 **Gemini CLI** | 通过 `GEMINI.md` 或 skill activate 引用 |
| ⚪ **任何 LLM Agent** | 直接把 SKILL.md 内容贴到 system prompt 里就行 |

核心原理：每个 skill 的 `SKILL.md` 就是一份详细的工作指南——告诉 AI "遇到这类任务该怎么做、按什么步骤、注意什么规则"。不管是 Claude、GPT、Gemini 还是其他模型，只要能读懂 markdown，就能按照指南执行。

---

## 📦 有什么

| 场景 | Skill | 产出 |
|------|-------|------|
| 📊 **做 Deck** | cathay-ppt-template | 国泰品牌 IC memo、pitch deck、客户 presentation（.pptx） |
| 📈 **建模型** | cathay-excel-template | 三张表模型、DCF、PE returns 分析（.xlsx） |
| 🎯 **Market Sizing** | market-sizing | 自下而上的 TAM/SAM/SOM，直接出 Excel |
| 🔗 **产业链分析** | chain-screener | 供应链 mapping + 流程图 + Excel 报告 |
| 🔍 **选股筛选** | stock-screener | 5 层 AI 主题筛选 |
| ⚖️ **相对估值** | stock-compare | 同行对比、相对强度分析 |
| 🧠 **深度研究** | equity-research | 7 个 AI agent 辩论式研究（多空对辩 + CIO 拍板） |

---

## 📁 目录结构

```
├── templates/
│   ├── cathay-ppt/          # PPT 模板 + Python 生成工具库
│   │   ├── assets/          # template.pptx（国泰品牌，12 种版式）
│   │   ├── lib/             # 文字引擎、slide 模板、QC 自动化、数据驱动生成
│   │   └── references/      # 生成规则、文字排版规范
│   └── cathay-excel/        # Excel 模型模板 + Python 工具库
│       ├── assets/          # template.xlsx（13 个 sheet）
│       ├── lib/             # 公式引擎、行号映射、模型填充、验证
│       └── docs/            # 设计文档
│
├── skills/
│   ├── market-sizing/       # TAM/SAM/SOM 分析 → Excel
│   ├── chain-screener/      # 产业链 mapping
│   ├── stock-screener/      # 主题选股
│   ├── stock-compare/       # 相对估值
│   └── equity-research/     # MoA 辩论式研究 + 估值框架
│
└── docs/
    └── setup.md             # API 配置说明
```

---

## 🔬 各 Skill 详细说明

### 📊 PPT 模板（cathay-ppt-template）

这个是我花了最多时间打磨的。核心解决一个痛点：**中文内容在 PPT 里排版溢出**。

- 12 种 PowerPoint 版式 + 16 种预设 slide 模板（标题页、内容页、对比、图表、SWOT、瀑布图、漏斗等）
- 专门写了 CJK 文字宽度计算引擎，中文再也不会溢出框外了
- 8 条 QC 规则 + 自动修复 pipeline（内容超出、字号不一致、布局重复等都会自动纠正）
- 品牌色：Maroon (#800000)、Gold (#E8B012)；字体：英文 Calibri、中文楷体

### 📈 Excel 模型（cathay-excel-template）

给 PE deal 用的标准财务模型：

- 13 个 sheet（Cover → 收入拆分 → COGS/OpEx → 三张表 → 运营资本 → 债务/CapEx → Returns & Sensitivity → DCF → Comps → Dashboard）
- 617 条预验证的 Excel 公式，全部互相链接
- row-map 系统：彻底杜绝行号偏移 bug（做过模型的都懂这个痛点）
- 10 项自动验证（BS 是否平、现金是否 tie-out、公式完整性等）
- 三种情景一键切换（Base / Upside / Downside）

### 🎯 Market Sizing

做 market sizing 最怕口径不一致。这个 skill 会强制你先确认边界（地理、产品、收入口径、时间范围），确认完才让你开始建模：

- 自下而上的方法论
- 供需平衡分析
- 8 个 sheet 的 Excel 输出，每个数字都有 audit trail
- 不允许静默 fallback——数据拿不到就报错，绝不瞎编

### 🧠 Equity Research（MoA 辩论）

这个比较有意思——7 个 AI agent 轮流发言，模拟真实研究团队的辩论过程：

1. 宏观策略师 → 2. 审计师 → 3. 行业专家 → 4. 多头 → 5. 空头 → 6. 多空互怼 → 7. CIO 最终拍板

支持 DCF、P/E-Growth、EV/EBITDA、NAV、FCF Yield 等估值范式。快速模式出 10-15 页 slides，深度模式出 25-40 页 + 配套 Excel。

---

## ⚠️ 注意事项

- 这些是 **AI agent 的 skill 定义文件**——教 AI 怎么完成任务的，不是独立命令行工具
- `chain-screener` 和 `stock-screener` 的 Python 脚本依赖外部量化系统（`QUANT_ROOT`），这里放的是参考实现
- `stock-screener` 完整运行需要一个跑着的 FastAPI 服务（量化系统的一部分）
- `stock-compare` 引用了外部的行情脚本和回测工具

---

## 🚀 怎么用

下面以 Claude Code 为例演示，其他平台类似——核心就是把 SKILL.md 放到 agent 能读到的地方。

### Step 1：装 Claude Code

```bash
# macOS / Linux
npm install -g @anthropic-ai/claude-code
```

或者直接用 Claude Code 桌面版 / VS Code 插件也行。

### Step 2：把 Skill 放到对应目录

```bash
# 克隆这个 repo
git clone https://github.com/harryhuang0719/Skills-for-Cathay.git
cd Skills-for-Cathay

# PPT 和 Excel 模板放到 Claude Code skills 目录
cp -r templates/cathay-ppt ~/.claude/skills/cathay-ppt-template
cp -r templates/cathay-excel ~/.claude/skills/cathay-excel-template

# 研究类 skill 也放进去
cp -r skills/equity-research ~/.claude/skills/equity-research
cp -r skills/market-sizing ~/.claude/skills/market-sizing
cp -r skills/chain-screener ~/.claude/skills/chain-screener
cp -r skills/stock-screener ~/.claude/skills/stock-screener
cp -r skills/stock-compare ~/.claude/skills/stock-compare
```

### Step 3：装 Python 依赖

```bash
pip install python-pptx openpyxl
```

### Step 4：配 API Key（按需）

PPT 和 Excel 模板 **不需要任何 API key**，装完直接能用。

其他 skill 按需配置（终端 export 或写到 `.env`）：

```bash
# 公司财务数据（chain-screener, stock-screener, equity-research 需要）
export FMP_API_KEY="你的key"
# 注册：https://financialmodelingprep.com/developer/docs/（免费 250 次/天）

# LLM 分析（market-sizing 需要）
export GEMINI_API_KEY="你的key"
# 注册：https://ai.google.dev/gemini-api/docs/api-key（Google AI Studio 里创建）

# LLM 深度分析（equity-research 需要）
export ANTHROPIC_API_KEY="你的key"
# 注册：https://console.anthropic.com/

# 中国/港股数据（chain-screener 可选）
export TUSHARE_TOKEN="你的token"
# 注册：https://tushare.pro/
```

### Step 5：开始用 ✨

打开 Claude Code，直接说人话就行：

```
# 做 PPT
"帮我做一个关于 XX 公司的 IC memo deck"
"做一个 5 页的 pitch deck，介绍我们的新基金"

# 建模型
"帮我建一个 XX 公司的三张表模型，用最近 4 年的历史数据"
"做一个 DCF，假设收入增速 20%，terminal growth 3%"

# Market Sizing
"帮我做一个中国工业机器人市场的 TAM sizing"

# 产业链
"分析一下光伏产业链，找出上下游相关上市公司"

# 深度研究
"用深度模式分析 NVDA，我要看多空双方的完整论点"
```

---

## 🔄 其他平台怎么用

### OpenAI Codex CLI
```bash
# 在项目根目录创建 AGENTS.md，引用 skill
echo "请参考 skills/market-sizing/SKILL.md 中的规则来执行 market sizing 任务" > AGENTS.md
codex "帮我做一个中国储能市场的 TAM sizing"
```

### Cursor / Windsurf
把 SKILL.md 加到项目 rules 里：
```bash
cp templates/cathay-ppt/SKILL.md .cursor/rules/cathay-ppt.md
```
然后在 Cursor 里直接说 "帮我做一个 IC memo deck"，它就会按规则来。

### 通用方式（任何 LLM）
最简单粗暴：直接把 SKILL.md 的内容贴到对话的 system prompt 里，然后正常聊就行。Python lib 需要本地有对应环境。

---

## 📚 更多细节

详细的 API 配置说明见 [docs/setup.md](docs/setup.md)。

每个 skill 目录下的 `SKILL.md` 有完整的技术文档，包括所有参数、输入输出格式、工作流程等。想深入了解某个 skill 直接去读对应的 SKILL.md 就行。

---

## 📬 联系我

有问题随时找我聊：

**微信：18918509837**（Harry）

用的过程中遇到问题、想加新场景、或者有什么改进建议，都欢迎直接微信找我。
