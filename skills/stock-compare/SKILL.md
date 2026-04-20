---
name: stock-compare
description: "Compare stocks vs peers, sector, or index. Activate when user asks: stock A vs B, how does X compare to sector/industry, is X better than Y, relative valuation, peer comparison, 对比, 比较, X和Y哪个更好, X跑赢行业了吗, 相对强弱, 估值对比, 行业对比. More specific than finance or quant-analysis for comparison tasks."
metadata: { "openclaw": { "emoji": "⚖️" } }
---

## ⛽ Pre-Flight Gate [R80][R93]

□ Step 0 — 时间感知: `session_status` [R30]
□ Step 1 — 双 ticker 数据到位:
  ```bash
  cd "机长进化论/scripts" && python3 -c "
  from captain_data_layer import get_quote, get_fundamentals
  for t in ['{TICKER_A}', '{TICKER_B}']:
      print(t, get_quote(t))
      print(t, get_fundamentals(t))
  "
  ```
□ Step 2 — KB 检索两个 ticker [R31]
□ Step 3 — 交叉验证: 估值指标 ≥2 源 [R32]

⛔ 任一 ticker 数据缺失 → 标注并降级（可比较可用部分）

# Stock Comparison Skill

当用户问股票对比类问题时，必须按以下流程执行，不能只用语言回答。

## 触发场景

- "NVDA vs AMD"
- "ZS和CRWD哪个更好"
- "X跑赢行业了吗"
- "X的估值和同行比怎么样"
- "X相对SPY/QQQ表现如何"
- "这个行业里哪只股票最强"

## 执行流程（必须按顺序）

### Step 1: 知识库检索（必须）
```bash
python3 "/Users/harryhuang/Algo Trading/knowledge-base/_scripts/retrieve.py" \
  --query "<ticker_a> <ticker_b> comparison peer 对比" --top_k 5
```

### Step 2: 获取实时行情对比
```bash
cd ~/.openclaw/workspace/skills/finance && source .venv/bin/activate
python3 scripts/market_quote.py <TICKER_A>
python3 scripts/market_quote.py <TICKER_B>
python3 scripts/market_quote.py <SECTOR_ETF>
```
- 同行业对比时，自动加入对应行业ETF作为基准（网安→HACK，黄金→GDX，国防→ITA，科技→QQQ）

### Step 3: 基本面对比（FMP API 直接调用）
```bash
# FMP_API_KEY from $QUANT_ROOT/.env
curl -s "https://financialmodelingprep.com/stable/profile?symbol=<TICKER_A>&apikey=$FMP_API_KEY" | python3 -m json.tool
curl -s "https://financialmodelingprep.com/stable/profile?symbol=<TICKER_B>&apikey=$FMP_API_KEY" | python3 -m json.tool
```
对比维度：PE/PS/PB、Revenue Growth、Gross Margin、FCF Margin、Debt/Equity

### Step 4: 回测相对表现（如果DuckDB有数据）
```bash
export QUANT_ROOT="/Users/harryhuang/Algo Trading/Quant Trading"
python3 $QUANT_ROOT/skills/backtest/scripts/run_backtest.py portfolio \
  '{"name":"compare","positions":[{"ticker":"<A>","weight":0.5},{"ticker":"<B>","weight":0.5}]}' \
  --start <3个月前> --end <今天>
```
同时跑各自单独回测，对比Alpha/Sharpe/最大回撤。

### Step 5: 新闻情绪对比
```bash
cd ~/.openclaw/workspace/skills/news-sentiment
python3 scripts/news_sentiment.py <TICKER_A> --days 7
python3 scripts/news_sentiment.py <TICKER_B> --days 7
```

## 输出格式

```
⚖️ <A> vs <B> 对比分析

📊 价格表现（近3个月）
- <A>: +X% | <B>: +Y% | 行业ETF: +Z%

💰 估值对比
| 指标 | <A> | <B> | 行业均值 |
|------|-----|-----|---------|
| PE   |     |     |         |
| PS   |     |     |         |
| 毛利率|    |     |         |

📈 技术面
- <A>: RSI=X, 趋势=...
- <B>: RSI=X, 趋势=...

📰 近期情绪
- <A>: 📈/➡️/📉 (score)
- <B>: 📈/➡️/📉 (score)

🎯 机长判断
- 相对强弱：<A>/<B> 更强，原因...
- 当前更值得关注：...
- 风险提示：...
```

## 行业ETF映射

| 行业 | ETF | 代表股 |
|------|-----|--------|
| 网络安全 | HACK | CRWD, ZS, PANW |
| 黄金/贵金属 | GDX | NEM, WPM, AEM |
| 国防 | ITA | LMT, NOC, RTX |
| 科技大盘 | QQQ | NVDA, AAPL, MSFT |
| 消费必需 | XLP | COST, WMT, PG |
| 大盘基准 | SPY | — |
