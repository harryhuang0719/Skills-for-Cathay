---
name: stock-screener
description: AI-powered thematic stock screener that validates investment theses and discovers relevant stocks using multi-layer analysis (thesis validation, discovery, quantitative screening, smart money detection, and AI ranking).
metadata:
  {
    "clawdbot":
      {
        "config":
          {
            "requiredEnv": ["SCREENER_ROOT"],
            "stateDirs": [".cache/screener"],
            "example": "export SCREENER_ROOT=\"/Users/harryhuang/Algo Trading/knowledge-base/screener_v2\"\n",
          },
      },
  }
---

# Stock Screener Skill

## ⛽ Pre-Flight Gate [R80][R93]

> 筛选前的环境验证。

□ Step 0 — 时间感知: `session_status` [R30]
□ Step 1 — FastAPI Server 健康检查:
  ```bash
  curl -sf --max-time 5 http://localhost:8000/health && echo "✅ Screener API OK" || echo "❌ Screener API down"
  ```
  不可用 → 提示: "请先启动 screener: cd $SCREENER_ROOT && uvicorn main:app --port 8000"
□ Step 2 — FMP API 可用 (screener 依赖):
  ```bash
  [ -n "$FMP_API_KEY" ] && echo "✅ FMP_API_KEY set" || echo "❌ FMP_API_KEY missing"
  ```
□ Step 3 — KB 补充: get_kb_context 获取主题相关历史材料 [R31]

⛔ FastAPI 不可用 → 中止，不编造筛选结果

This skill provides **AI-powered thematic stock screening** using a 5-layer architecture:

- **Layer 0**: Thesis validation (scores investment thesis 0-10)
- **Layer 1**: Discovery (finds relevant tickers by sector/keywords)
- **Layer 2**: Quantitative screening (valuation, growth, momentum, quality)
- **Layer 3**: Smart money detection (volume anomalies, options flow)
- **Layer 4**: AI ranking (LLM-generated investment summaries)

## Knowledge Base Integration (MANDATORY)

Before generating any analysis output, ALWAYS check the local knowledge base for relevant materials:

1. Read `/Users/harryhuang/Algo Trading/knowledge-base/index.jsonl`
2. Filter entries by matching tickers, tags, or keywords related to the stock/sector being analyzed
3. Load and incorporate relevant files (expert interviews, research reports, meeting notes)
4. Cite knowledge base sources in the final report: `根据知识库材料（expert_interviews/xxx.docx）...`

This step is non-optional — local expert interviews and research reports take priority over generic LLM knowledge.

## When to use

Use this skill when the user asks:

- "帮我找**相关的股票" / "Find stocks related to **"
- "筛选**主题的投资机会" / "Screen for ** theme opportunities"
- "哪些股票受益于**" / "Which stocks benefit from **"
- "存储芯片短缺有什么投资机会" / "Investment opportunities in memory chip shortage"
- "AI芯片相关的股票" / "AI chip related stocks"

**Important**: This skill is for **thematic screening** (e.g., "AI chips", "memory shortage"). For individual stock analysis, use the `quant-analysis` skill instead.

---

## Commands

### Thematic stock screening

Validates an investment thesis and finds the best stock opportunities.

**Usage**:

```bash
python3 ~/.openclaw/workspace/skills/stock-screener/scripts/screen_thesis.py "THESIS" [--top N] [--aggression LEVEL]
```

**Examples**:

```bash
# Screen for memory chip shortage opportunities
python3 ~/.openclaw/workspace/skills/stock-screener/scripts/screen_thesis.py "存储芯片短缺"

# Find top 15 AI chip stocks
python3 ~/.openclaw/workspace/skills/stock-screener/scripts/screen_thesis.py "AI芯片" --top 15

# Aggressive screening for EV battery theme
python3 ~/.openclaw/workspace/skills/stock-screener/scripts/screen_thesis.py "电动车电池" --aggression 1.5
```

**Parameters**:

- `THESIS` (required): Investment thesis in Chinese or English
- `--top` (optional): Number of top picks to return (default: 10, max: 30)
- `--aggression` (optional): Factor weight adjustment (0.5-2.0, default: auto-calculated from thesis strength)

**Output**:

- Thesis validation score (0-10)
- Number of candidates discovered
- Top N stock picks with:
  - Ticker, name, price, market cap
  - Composite score (0-100)
  - Factor breakdown (valuation, growth, momentum, quality, theme purity)
  - Smart money signals (volume anomalies, options flow)
  - AI-generated investment summary

**Execution time**: 1-3 minutes

---

## Output expectations (how to present results to user)

### For screen_thesis:

1. **Immediately after starting**: Tell user "⏳ 正在分析「{THESIS}」主题，预计需要1-3分钟..."
2. **After thesis validation (Layer 0)**:
   - If rejected (verdict = "rejected"):
     ```
     ❌ 观点验证未通过 ({score}/10)
     反驳证据: {evidence_against 列表}
     建议: 这个观点目前缺乏足够的数据支撑。你可以重新表述或等待更多数据确认。
     ```
   - If passed: show `{verdict_emoji} 观点验证: {verdict_cn} ({score}/10)` + top 3 evidence_for + any evidence_against
   - verdict translations: strong → ✅ 强支撑 | moderate → ⚠️ 中等支撑 | weak → ⚡ 弱支撑
3. **After discovery (Layer 1)**:
   - Show: "🔍 发现 {count} 个相关候选股票 | 模式: {aggression_cn}"
   - aggression translations: high → 激进模式 | medium → 均衡模式 | low → 保守模式
4. **After completion**, present results in **3-tier table format**:

```
🏆 Tier 1 — 最佳标的
| 代码 | 公司 | 价格 | 今日涨跌 | 综合评分 | 成长 | 动量 | 估值 | 质量 |
|------|------|------|----------|----------|------|------|------|------|
| MU   | Micron | $95.20 | +2.3% | 8.7/10 | ★★★★★ | ★★★☆☆ | ★★★★☆ | ★★★★☆ |
🔥 Smart Money: {异常信号描述，如有}
💡 {AI summary}

🥈 Tier 2 — 次选标的
| 代码 | 公司 | 价格 | 今日涨跌 | 综合评分 | 成长 | 动量 | 估值 | 质量 |
(2 tickers，同上格式)

🥉 Tier 3 — 关注标的
| 代码 | 公司 | 价格 | 今日涨跌 | 综合评分 | 成长 | 动量 | 估值 | 质量 |
(3 tickers，同上格式)

📝 结论：{2-3句：主题逻辑、最佳切入点、主要风险}
⏱ 筛选耗时: {execution_time}秒
💡 要我对其中某个标的跑完整的深度分析吗？
```

   - Tier 1: rank 1 | Tier 2: rank 2-3 | Tier 3: rank 4-6
   - 价格和涨跌幅从返回数据的 `price` / `change_pct` 字段读取
   - Smart Money 行仅在 `smart_money_score >= 6` 时显示
   - Add disclaimer at the end

---

## Natural language understanding

The agent should recognize these patterns and map them to commands:

| User says                                  | Command to run                                                                                               |
| ------------------------------------------ | ------------------------------------------------------------------------------------------------------------ |
| "帮我找存储芯片短缺相关的股票"             | `python3 ~/.openclaw/workspace/skills/stock-screener/scripts/screen_thesis.py "存储芯片短缺"`                |
| "AI芯片有什么投资机会"                     | `python3 ~/.openclaw/workspace/skills/stock-screener/scripts/screen_thesis.py "AI芯片"`                      |
| "筛选电动车电池主题，要激进一点"           | `python3 ~/.openclaw/workspace/skills/stock-screener/scripts/screen_thesis.py "电动车电池" --aggression 1.5` |
| "Find stocks benefiting from cloud growth" | `python3 ~/.openclaw/workspace/skills/stock-screener/scripts/screen_thesis.py "cloud computing growth"`      |

**Important**: The agent should:

- Accept both Chinese and English theses
- Infer aggression level from context:
  - "激进" / "aggressive" → 1.5
  - "保守" / "conservative" → 0.7
  - Default → auto (based on thesis strength)
- Adjust top N based on user request:
  - "前5个" → --top 5
  - "多一些" → --top 20
  - Default → 10

---

## Technical requirements

### Environment setup

- **Required environment variable**: `SCREENER_ROOT`
  - Should point to the screener v2 directory
  - Example: `/Users/harryhuang/Algo Trading/knowledge-base/screener_v2`
- **Working directory**: All scripts should execute with `cwd=$SCREENER_ROOT`
- **FastAPI server**: The screener runs as a FastAPI service on `http://localhost:8000`

### Starting the service

Before using the skill, ensure the FastAPI server is running:

```bash
cd "$SCREENER_ROOT"
source venv/bin/activate
uvicorn main:app --host 0.0.0.0 --port 8000
```

The skill script will check if the service is running and prompt the user to start it if needed.

### Dependencies

The screener system has its own virtual environment with all dependencies. The skill script is a thin wrapper that calls the FastAPI endpoint.

### Error handling

1. **If SCREENER_ROOT not set**: Tell user to configure it
2. **If service not running**: "请先启动 Screener 服务：`cd $SCREENER_ROOT && source venv/bin/activate && uvicorn main:app --port 8000`"
3. **If thesis score < 5**: "该主题评分较低 ({score}/10)，建议重新表述或选择更具体的投资主题"
4. **If no candidates found**: "未找到相关股票，请尝试更广泛的主题或不同的关键词"
5. **If API error**: Check logs and suggest verifying API keys in `.env` file

---

## Prediction Logging (MANDATORY)

After every `screen_thesis.py` run, log the top 3 picks as predictions:

```bash
# For each of the top 3 picks:
python3 ~/.openclaw/workspace/skills/prediction-tracker/scripts/log_prediction.py \
  --skill "stock-screener" \
  --ticker {TICKER} \
  --direction long \
  --entry-price {CURRENT_PRICE from screener output} \
  --confidence {high if composite_score >= 80, medium if >= 60, low otherwise} \
  --timeframe 30d \
  --thesis "{THESIS} — composite score {SCORE}/100" \
  --tags "{theme_tags}"
```

This enables tracking which screener themes and factor profiles produce the best hit rates over time.

---

## Safety and best practices

1. **Rate limiting**: Maximum 2 concurrent screening tasks
2. **Timeout**: If a command takes > 5 minutes, consider it failed
3. **Cache results**: If user asks about same thesis within 30 minutes, reuse results
4. **Privacy**: Never share API keys or sensitive configuration in responses
5. **Thesis quality**: Encourage specific, actionable theses over vague ones

---

## Disclaimer (always include when giving investment insights)

After providing screening results, always add:

```
⚠️ 免责声明：本筛选结果仅供参考，不构成投资建议。投资有风险，决策需谨慎。
⚠️ Disclaimer: This screening is for informational purposes only and does not constitute investment advice. Investing involves risk.
```

---

## v2.0 升级: Composite Score 公式 + Thesis 阈值 + Backtest 验证 (2026-03-23)

### 1. Composite Score 标准公式

解决原 0-100 分无明确公式的问题:

```
composite_score = (
    valuation_score  × 0.25 +
    growth_score     × 0.30 +
    momentum_score   × 0.25 +
    quality_score    × 0.20
) × 100
```

**各子分标准化 (0-1)**:

| 维度 | 指标 | 标准化方法 |
|------|------|-----------|
| valuation_score | Forward PE, PS, PB | `1 - percentile_rank(PE, sector)` (PE 越低越好) |
| growth_score | Revenue Growth YoY, EPS Growth YoY | `percentile_rank((rev_g + eps_g) / 2, universe)` |
| momentum_score | 3M Price Return | `percentile_rank(return_3m, sector)` |
| quality_score | ROE, Net Margin, Debt/Equity | `(percentile_rank(ROE) + percentile_rank(margin) + (1 - percentile_rank(debt))) / 3` |

**percentile_rank**: 在同 sector 或 universe 中的百分位排名 (0=最差, 1=最好)

### 2. Thesis Validation 阈值定义

| Score | Verdict | Emoji | 说明 |
|-------|---------|-------|------|
| ≥ 8.0 | **strong** | ✅ 强支撑 | 多维度数据支持，可直接进入 screener |
| 5.0-7.9 | **moderate** | ⚠️ 中等支撑 | 部分维度支持，需进一步验证 |
| 3.0-4.9 | **weak** | ⚡ 弱支撑 | 数据支持不足，谨慎使用 |
| < 3.0 | **rejected** | ❌ 未通过 | 与事实严重不符，跳过该 thesis |

**Rejected 处理**: 显示 evidence_against 列表，建议修改 thesis 方向

### 3. Backtest 验证层

Top 5 选股结果自动接入 backtest skill 进行历史验证:

```bash
# 自动运行: top 5 picks 过去 60 天表现
python3 "/Users/harryhuang/Algo Trading/Quant Trading/skills/backtest/scripts/run_backtest.py" portfolio \
  '[{"ticker":"TOP1","weight":0.2},{"ticker":"TOP2","weight":0.2},...,{"ticker":"TOP5","weight":0.2}]' \
  --start {60天前} --end {today}
```

**输出追加**:
```
📊 历史验证 (60d backtest):
  组合收益: +8.3% vs SPY +4.2% (α = +4.1%)
  最大回撤: -6.2%
  Sharpe: 1.45
  Top performer: {TICKER} +15.2%
  Worst performer: {TICKER} -3.1%
```

**规则**:
- Backtest α < 0 → 追加 "⚠️ 历史验证未通过: 该 thesis 在过去 60 天跑输基准"
- Backtest max_drawdown > 15% → 追加 "⚠️ 高回撤风险"
- Backtest 数据不足 (coverage < 80%) → 标注 "数据覆盖不足，仅供参考"

---

## 📤 Output & Distribution [R96]

| 渠道 | 格式 | 路由 |
|------|------|------|
| 终端 | 3-tier 排名表 + 因子分解 | 默认 |
| Dashboard | JSON → /api/skills/run | 自动 |
| 飞书 | Top picks → 选股群 | oc_9bc28b67460d8bb3cc063d7f44ddb792 |
| Discord | Markdown ranking | reply tool |

## 🔗 Cross-Skill Hooks [R95]

**上游依赖**: FastAPI screener server, FMP API, knowledge-retrieval
**下游触发**:
- prediction-tracker [R95]: Top 3 picks 必须记录
  ```bash
  python3 ~/.openclaw/workspace/skills/prediction-tracker/scripts/log_prediction.py \
    --skill stock-screener --ticker {TICKER} --direction long \
    --entry-price {PRICE} --confidence {high|medium|low} \
    --timeframe 30d --thesis "{THEME}: score {SCORE}"
  ```
- pipeline /sector-scan: 作为 Step 3 量化筛选
