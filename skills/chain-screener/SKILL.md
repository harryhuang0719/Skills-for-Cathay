---
name: chain-screener
description: "产业链 Screener — 输入投资主题（如'SpaceX产业链'、'AI算力'、'人形机器人'），自动拆解上游/中游/下游产业链，全球搜索相关上市公司（美股/港股/A股），拉取财务数据，生成投资评分、Mermaid流程图和Excel报表。Triggers on: 产业链, 供应链, 上下游, 概念股, 受益股, 相关标的, industry chain, supply chain."
metadata:
  {
    "clawdbot":
      {
        "config":
          {
            "requiredEnv": ["QUANT_ROOT"],
            "example": "export QUANT_ROOT=\"/Users/harryhuang/Algo Trading/Quant Trading\"\n",
          },
      },
  }
---

# 产业链 Screener Skill

## ⛽ Pre-Flight Gate [R80][R93]

> 产业链筛选前的数据源验证。

□ Step 0 — 时间感知: `session_status` [R30]
□ Step 1 — 环境变量:
  ```bash
  [ -n "$QUANT_ROOT" ] && [ -n "$FMP_API_KEY" ] && echo "✅ ENV OK" || echo "❌ Missing QUANT_ROOT or FMP_API_KEY"
  ```
□ Step 2 — 长时间运行提醒: 本 skill 需 30-90 秒，必须用 `sessions_spawn` (timeout=300)
□ Step 3 — KB 补充: 检索主题相关知识库材料 [R31]

⛔ FMP 不可用 → 输出标注 "⚠️ 部分数据缺失"（可继续但质量降级）

给定一个投资主题，在 30-60 秒内完成：

1. **LLM 产业链拆解** — 上游→中游→下游，每层 2-5 个子环节
2. **全球公司搜索验证** — FMP（美股/港股）+ Tushare（A股），自动纠正错误代码
3. **投资评分** — 100分制（题材纯度30 + 竞争地位25 + 估值20 + 成长15 + 流动性10）
4. **Mermaid 流程图** — 可直接粘贴到 mermaid.live
5. **Excel 报表** — 3 sheets：产业链全景 / Top投资标的 / 流程图代码

## When to use

- "分析 SpaceX 产业链" / "SpaceX supply chain"
- "AI算力有哪些投资标的"
- "人形机器人上下游"
- "低空经济概念股"
- "固态电池受益股"
- "帮我找 xxx 相关的上市公司"

## Command

**Execution time**: 30-90 秒 — **ALWAYS use `sessions_spawn` to avoid timeouts.**

### How to invoke

1. Tell the user: "⏳ 正在分析「{THEME}」产业链，完成后会通知你（约30-60秒）..."
2. Call `sessions_spawn` with the task below:

```
Task template for sessions_spawn:
---
Run 产业链 Screener for theme: {THEME}

Steps:
1. Run: python3 $QUANT_ROOT/skills/chain-screener/scripts/run_chain_screener.py "{THEME}" --output $QUANT_ROOT/screener_output
   - cwd: $QUANT_ROOT
   - Use exec with timeout=300 and yieldMs=60000
2. Parse the output line starting with "SUMMARY_JSON:" to get the result dict.
3. Send results to user via feishu DM (target: user:ou_393cc8e9f76de2380dd05213f578cf78):
   - Theme + investment logic (1 sentence)
   - Top 5 picks table: Rank | Company | Ticker | Market | Score
   - Send the Excel file: message(action=send, channel=feishu, filePath=<files.excel from summary>)
   - Send the PDF report: message(action=send, channel=feishu, filePath=<files.pdf from summary>)
   - Mermaid code block (copy from files.mermaid)
4. Always append: ⚠️ 免责声明：本分析仅供参考，不构成投资建议。投资有风险，决策需谨慎。
---
```

**Parameters**:

- `THEME` (required): 投资主题，中英文均可
  - 中文示例: `SpaceX产业链`, `AI算力`, `人形机器人`, `低空经济`, `固态电池`
  - 英文示例: `humanoid robots`, `AI data center`, `commercial space`

## Output

Script prints `SUMMARY_JSON:{...}` on the last line with:

```json
{
  "theme": "SpaceX产业链",
  "total_time": "45.2s",
  "total_companies": 48,
  "verified": 31,
  "markets": { "US": 18, "HK": 5, "A": 12, "UNLISTED": 13 },
  "top_5": [
    { "name": "SpaceX", "ticker": "未上市", "score": 82 },
    { "name": "Rocket Lab", "ticker": "RKLB", "score": 71 }
  ],
  "files": {
    "excel": "/path/to/SpaceX产业链_20260223120000.xlsx",
    "mermaid": "/path/to/SpaceX产业链_20260223120000.mermaid",
    "json": "/path/to/SpaceX产业链_20260223120000.json",
    "pdf": "/path/to/SpaceX产业链_20260223120000.pdf"
  }
}
```

## Scoring v2 (2026-04-19)

评分公式已升级，加入 forward-looking 和 peer-relative 维度：

| 维度 | 满分 | 数据源 |
|------|------|--------|
| 题材纯度 | 25 | LLM 标注 (★-★★★★★) |
| 竞争地位 | 20 | LLM 标注 (龙头/二线/新进入者) |
| 估值 | 20 | Forward PE (优先) + trailing PE, peer-relative bonus |
| 成长性 | 15 | 历史营收增速 + Forward EPS growth bonus |
| 流动性 | 8 | 市值分档 |
| 分析师覆盖 | 7 | FMP analyst estimates (numAnalystsEps) |
| Insider 信号 | 5 | FMP insider trading buy/sell ratio |
| 叙事评分 | 10 | narrative_scorer composite |
| 期权信号 | 5 | options_sentiment PCR |

数据来源：US 标的通过 `captain_data_layer.get_analyst_estimates()` 和 `get_insider()` 自动拉取 forward data。CN/HK 标的仍用基础评分。

归一化：原始分 ÷ 1.15 → 0-100 分。

## Natural language mapping

| User says                         | THEME to pass      |
| --------------------------------- | ------------------ |
| "SpaceX产业链" / "商业航天"       | `SpaceX产业链`     |
| "AI算力" / "AI芯片" / "GPU"       | `AI算力产业链`     |
| "人形机器人" / "humanoid"         | `人形机器人产业链` |
| "低空经济" / "eVTOL" / "飞行汽车" | `低空经济产业链`   |
| "固态电池"                        | `固态电池产业链`   |
| "光伏" / "solar"                  | `光伏产业链`       |

## Error handling

- **QUANT_ROOT not set**: Tell user `export QUANT_ROOT="/Users/harryhuang/Algo Trading/Quant Trading"`
- **LLM JSON parse error**: LLM returned malformed JSON — retry once with `--retry` flag (not yet implemented; ask user to retry)
- **FMP/Tushare API error**: Partial results still generated; verified count will be lower than total

---

⚠️ 免责声明：本分析仅供参考，不构成投资建议。投资有风险，决策需谨慎。

## 📤 Output & Distribution [R96]

| 渠道 | 格式 | 路由 |
|------|------|------|
| 终端 | 产业链图 + 投资评分 + Mermaid 图 | 默认 |
| Dashboard | JSON + Excel → /api/skills/run | 自动 |
| 飞书 | Excel + PDF → 选股群 | oc_9bc28b67460d8bb3cc063d7f44ddb792 |
| Discord | Markdown 产业链摘要 (2000 char) | reply tool |

输出: "数据: FMP {date}, Tushare {date}, KB {N} docs"
