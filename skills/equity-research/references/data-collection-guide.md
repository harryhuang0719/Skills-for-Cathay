# Data Collection Guide

## Environment Setup

All API keys are in `~/Algo Trading/Quant Trading/.env`. Load with:
```python
import os
from dotenv import load_dotenv
for p in [
    os.path.expanduser("~/Algo Trading/Quant Trading/.env"),
    os.path.expanduser("~/Algo Trading/Quant Trading/configs/.env"),
    os.path.expanduser("~/.openclaw/.env"),
]:
    if os.path.exists(p):
        load_dotenv(p)
```

## Data Routing by Market

| Market | Primary | Secondary | Tertiary |
|--------|---------|-----------|----------|
| US stocks | FMP API + Seeking Alpha (RapidAPI) | Tavily + Brave (news) | SEC EDGAR, FRED (macro) |
| HK stocks | Tavily + Brave | Tushare (HK daily) | DuckDB EOD |
| A-shares | Tavily + Brave | Tushare (A-share daily) | DuckDB EOD |

---

## 1. FMP API (Primary Financial Data)

```python
FMP_KEY = os.getenv('FMP_API_KEY', '')
BASE = "https://financialmodelingprep.com/stable"

def fmp(endpoint, **params):
    params['apikey'] = FMP_KEY
    r = requests.get(f"{BASE}/{endpoint}", params=params)
    r.raise_for_status()
    return r.json()
```

### Quick Mode Endpoints (7 calls)

**注意**: 即使Quick模式也必须拉季度数据（铁律：季度数据优先于年度数据）。

| Endpoint | Purpose |
|----------|---------|
| `profile?symbol={T}` | Sector, industry, description, market cap, CEO |
| `income-statement?symbol={T}&period=quarter&limit=6` | **季度**收入、利润率、EPS趋势（核心） |
| `income-statement?symbol={T}&period=annual&limit=3` | 年度汇总（辅助，不能替代季度） |
| `key-metrics?symbol={T}&period=quarter&limit=6` | 季度估值倍数、盈利能力 |
| `earnings-surprises?symbol={T}` | Beat/miss历史 |
| `analyst-estimates?symbol={T}&limit=4` | Consensus estimates (next 2Y) |
| `stock-price-change?symbol={T}` | 1M/3M/6M/1Y/YTD returns |

### Deep Mode Additional Endpoints (15+ calls)

**注意**: Deep模式必须同时拉季度+年度数据。季度数据是核心，年度数据是辅助汇总。

| Endpoint | Purpose |
|----------|---------|
| `income-statement?symbol={T}&period=quarter&limit=8` | **季度**IS趋势（核心，见SKILL.md铁律） |
| `income-statement?symbol={T}&period=annual&limit=5` | 年度IS汇总 |
| `balance-sheet-statement?symbol={T}&period=quarter&limit=2` | 最新季度BS |
| `balance-sheet-statement?symbol={T}&period=annual&limit=5` | BS 5Y趋势 |
| `cash-flow-statement?symbol={T}&period=quarter&limit=8` | **季度**CF趋势（核心） |
| `cash-flow-statement?symbol={T}&period=annual&limit=5` | CF 5Y汇总 |
| `revenue-product-segmentation?symbol={T}&period=quarter` | **季度**收入分拆 |
| `revenue-geographic-segmentation?symbol={T}&period=annual` | 地区收入分布 |
| `ratios?symbol={T}&period=annual&limit=5` | Full ratio suite |
| `enterprise-values?symbol={T}&period=annual&limit=5` | EV, EV/EBITDA |
| `financial-growth?symbol={T}&period=annual` | Growth rates |
| `historical-price-full?symbol={T}` | 3Y daily OHLCV |
| `earnings-surprises?symbol={T}` | Beat/miss history |
| `analyst-stock-recommendations?symbol={T}` | Individual analyst recs |
| `stock-peers?symbol={T}` | Comparable companies |
| `sec-filings-search?symbol={T}&type=10-K&limit=3` | SEC filing links |
| `stock-news?symbol={T}&limit=20` | Company news |

---

## 2. Seeking Alpha via RapidAPI (Analyst Data)

```python
RAPID_KEY = os.getenv('RAPIDAPI_KEY', '')
SA_HEADERS = {
    "X-RapidAPI-Key": RAPID_KEY,
    "X-RapidAPI-Host": "seeking-alpha.p.rapidapi.com"
}
SA_BASE = "https://seeking-alpha.p.rapidapi.com"
```

### Key Endpoints

| Endpoint | Purpose |
|----------|---------|
| `symbols/get-analyst-price-target?symbol={T}` | Consensus PT (high/low/mean/median, # analysts) |
| `symbols/get-summary?symbols={T}` | Forward PE, EPS, div yield |
| `symbols/get-estimates?symbol={T}&data_type=quarterly` | Quarterly revenue/EPS estimates |
| `symbols/get-fundamentals?symbol={T}` | Analyst rating time series |
| `symbols/get-metric-grades?symbol={T}` | SA quant grades (1-13 scale) |
| `symbols/get-sector-metrics?symbol={T}` | Industry median benchmarks |
| `transcripts/v2/list?id={T}` | Earnings call transcript list |
| `transcripts/v2/get-details?id={transcript_id}` | Full transcript text |

**Budget**: ~1,050 calls/month for daily scan. Be efficient.

---

## 3. Tavily API (AI-Powered News Search)

```python
TAVILY_KEY = os.getenv('TAVILY_API_KEY', '')
# 10 backup keys available: TAVILY_API_KEY_1 through TAVILY_API_KEY_10

def tavily_search(query, days=30, max_results=8):
    r = requests.post("https://api.tavily.com/search", json={
        "api_key": TAVILY_KEY,
        "query": query,
        "search_depth": "advanced",
        "topic": "finance",
        "max_results": max_results,
        "days": days
    }, timeout=15)
    return r.json().get("results", [])
```

### Search Queries for Equity Research

```python
# 通用查询（所有公司必须执行）
queries = [
    f"{COMPANY} latest earnings results {YEAR}",
    f"{TICKER} analyst target price consensus",
    f"{COMPANY} {INDUSTRY} outlook growth drivers",  # 用FMP profile返回的industry字段
    f"{COMPANY} competitive landscape market share",
]

# 行业专属查询（根据sector/industry动态生成，示例）:
# - Semiconductor: f"{COMPANY} HBM AI data center demand"
# - Oil & Gas: f"{COMPANY} production guidance OPEC outlook"
# - Financials: f"{COMPANY} NIM credit quality loan growth"
# - Healthcare: f"{COMPANY} pipeline FDA approval catalyst"
# - Consumer: f"{COMPANY} same-store sales consumer spending"
# - Mining: f"{COMPANY} commodity price production cost"
```

**Key rotation**: If primary key rate-limited, cycle through backup keys.
**Content fetching**: Use `include_raw_content: True` for full article text.

---

## 4. Brave Search API (News + Web)

```python
BRAVE_KEY = os.getenv('BRAVE_API_KEY', '')

def brave_search(query, count=5):
    r = requests.get("https://api.search.brave.com/res/v1/web/search",
        params={"q": query, "count": count},
        headers={"X-Subscription-Token": BRAVE_KEY}, timeout=10)
    return r.json().get("web", {}).get("results", [])

def brave_news(query, count=10, freshness="p7d"):
    r = requests.get("https://api.search.brave.com/res/v1/news/search",
        params={"q": query, "count": count, "freshness": freshness},
        headers={"X-Subscription-Token": BRAVE_KEY}, timeout=10)
    return r.json().get("results", [])
```

**Use for**: Analyst consensus, recent news, industry reports. Good fallback when Tavily quota exhausted.

---

## 5. SEC EDGAR (Deep Mode, US Stocks)

Via FMP endpoint: `sec-filings-search?symbol={T}&type=10-K&limit=3`

For direct SEC access:
```python
# CIK lookup
r = requests.get(f"https://efts.sec.gov/LATEST/search-index?q={TICKER}",
    headers={"User-Agent": "harryhhx@gmail.com"})
```

Filing types: 10-K (annual), 10-Q (quarterly), 8-K (current events), DEF 14A (proxy)

---

## 6. FRED (Macro Data, Deep Mode)

```python
FRED_KEY = os.getenv('FRED_API_KEY', '')
# Or use fredapi library
from fredapi import Fred
fred = Fred(api_key=FRED_KEY)
```

Key series for equity research context:
- `DFF` — Fed Funds Rate
- `DGS10` / `DGS2` — Treasury 10Y/2Y
- `T10Y2Y` — Yield curve spread
- `CPIAUCSL` — CPI
- `UNRATE` — Unemployment
- `VIXCLS` — VIX

---

## 7. DuckDB (Historical Pricing)

```python
import duckdb
DB_PATH = os.path.expanduser("~/Algo Trading/Quant Trading/data/quant.duckdb")
con = duckdb.connect(DB_PATH, read_only=True)
```

- 633MB, 12K tickers (5,292 CN, 4,391 US, 2,322 HK)
- Tables: `tickers`, `market_data`, `fundamentals`, `news`, `sec_filings_index`
- Use for HK/A-share historical pricing when FMP unavailable

---

## 8. Tushare (A-Share + HK Data)

```python
TUSHARE_TOKEN = os.getenv('TUSHARE_TOKEN', '')
# Gateway: http://1w1a.xiximiao.com/dataapi (积分网关)
# HK Gateway: http://hk_daily.xiximiao.com/dataapi
```

- A-share daily OHLCV, fundamentals, financials
- HK daily prices
- Use for non-US coverage

---

## 9. News Aggregation Strategy

**Fallback chain** (use multiple sources concurrently):
1. **Tavily** (primary) — advanced semantic search, finance topic
2. **Brave** (secondary) — web + news search
3. **FMP News** — `stock-news?symbol={T}&limit=20`
4. **Google News RSS** — `https://news.google.com/rss/search?q={COMPANY}+stock`

**Source credibility weighting**:
- Tier 1 (1.0): Bloomberg, Reuters, WSJ, FT
- Tier 2 (0.7-0.8): CNBC, Barron's, Seeking Alpha, Yahoo Finance
- Tier 3 (0.4-0.5): Benzinga, Motley Fool, MarketWatch

---

## 10. Data Completeness Check (Checkpoint 1)

Before proceeding to analysis, present this summary to user:

```
数据收集完成 — {TICKER} ({COMPANY_NAME})

■ 公司概况: {sector} / {industry} / 市值 {market_cap}
■ 财务数据: {n_years}年历史数据 (FMP)
■ 估值指标: P/E {pe}, EV/EBITDA {ev_ebitda}, FCF Yield {fcf_yield}
■ 分析师共识: 目标价 ${target} ({n_analysts} analysts) [来源: SA/Brave]
■ 收入分拆: {segment_summary} (FMP segments)
■ 最新新闻: {n_articles}条 [来源: Tavily/Brave/FMP]
■ 知识库: {n_kb_entries}条相关材料
■ 行业范式: {paradigm_name} → 估值方法: {methods}

是否继续分析？如有缺失数据请告知。
```
