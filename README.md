# Skills for Cathay

AI-powered skill toolkit for PE/VC primary market work. Covers deal materials generation, financial modeling, market sizing, and sector analysis.

## Capabilities

| Category | Skills | Output |
|----------|--------|--------|
| **Deck Generation** | cathay-ppt-template | Cathay-branded IC memos, pitch decks, client presentations (.pptx) — 12 layouts, 16 slide templates |
| **Financial Modeling** | cathay-excel-template | 3-statement models, DCF, PE returns analysis (.xlsx) |
| **Market Sizing** | market-sizing | Bottom-up TAM/SAM/SOM with Excel output |
| **Sector Analysis** | chain-screener | Supply chain mapping + Mermaid diagrams + Excel |
| **Stock Screening** | stock-screener | Thematic AI-powered screening (5-layer architecture) |
| **Relative Valuation** | stock-compare | Peer comparison, relative strength analysis |
| **Equity Research** | equity-research | 7-agent MoA debate, valuation paradigms, slide structures |

## Directory Structure

```
├── templates/
│   ├── cathay-ppt/          # PPT template + Python generation libs
│   │   ├── assets/          # template.pptx (Cathay brand, 12 layouts)
│   │   ├── lib/             # text_engine, slide_templates, qc_automation, data_driven
│   │   └── references/      # generation rules, text fitting specs
│   └── cathay-excel/        # Excel model template + Python libs
│       ├── assets/          # template.xlsx (13 sheets)
│       ├── lib/             # formula_engine, row_map, model_populator, validate_model
│       └── docs/            # design specs
│
├── skills/
│   ├── market-sizing/       # TAM/SAM/SOM analysis → Excel
│   ├── chain-screener/      # Industry chain mapping
│   ├── stock-screener/      # Thematic stock screening
│   ├── stock-compare/       # Relative valuation
│   └── equity-research/     # MoA debate + valuation frameworks
│
└── docs/
    └── setup.md             # API keys & environment setup
```

## Quick Start

### As Claude Code Skills

Copy skill directories into your Claude Code skills folder:

```bash
# PPT/Excel templates
cp -r templates/cathay-ppt ~/.claude/skills/cathay-ppt-template
cp -r templates/cathay-excel ~/.claude/skills/cathay-excel-template

# Analysis skills (for OpenClaw agent)
cp -r skills/market-sizing ~/.openclaw/workspace/skills/
cp -r skills/chain-screener ~/.openclaw/workspace/skills/
```

### Python Dependencies

```bash
pip install python-pptx openpyxl
```

## API Requirements

See [docs/setup.md](docs/setup.md) for detailed setup instructions.

| API | Required By | Purpose | Free Tier |
|-----|-------------|---------|-----------|
| FMP (Financial Modeling Prep) | chain-screener, stock-screener, equity-research | Company financials, quotes, peers | 250 calls/day |
| Tushare Pro | chain-screener (optional) | China/HK stock data | Yes (basic) |
| Google Gemini | market-sizing | LLM analysis for TAM modeling | Yes |
| Anthropic Claude | equity-research, market-sizing | Multi-agent debate, deep analysis | Pay-as-you-go |
| python-pptx | cathay-ppt | PPT generation | Open source |
| openpyxl | cathay-excel, market-sizing | Excel generation | Open source |

**No API needed:** cathay-ppt-template, cathay-excel-template (pure local generation)

## Key Features

### Cathay PPT Template
- 12 PowerPoint layouts + 16 pre-built slide templates (T1-T16): title, content, comparison, charts, SWOT, waterfall, funnel
- CJK-aware text fitting engine (Chinese character width handling)
- 8-rule QC automation with auto-fix pipeline
- Brand colors: Maroon (#800000), Gold (#E8B012), fonts: Calibri + KaiTi

### Cathay Excel Template
- 13-sheet PE financial model (IS, BS, CF, DCF, Comps, Returns)
- 617 pre-validated Excel formulas
- Row-map system eliminates row-offset bugs
- 10-point model validation (BS balance, cash tie-out, formula integrity)
- Scenario toggle (Base/Upside/Downside)

### Market Sizing
- Bottom-up methodology with mandatory boundary confirmation
- Supply-demand balance analysis
- 8-sheet Excel output with audit trail
- Execution integrity: no silent fallbacks, formula realization gates

### Equity Research (MoA Debate)
- 7-agent Mixture-of-Agents: Macro Strategist → Forensic Accountant → Industry Expert → Bull → Bear → Rebuttals → CIO
- Valuation paradigms: DCF, P/E-Growth, EV/EBITDA, NAV, FCF Yield
- Quick mode (10-15 slides) and Deep mode (25-40 slides + Excel)

## Important Notes

- **Skills are Claude Code / OpenClaw skill definitions** — they describe how AI agents perform tasks, not standalone CLI tools.
- `chain-screener` and `stock-screener` scripts depend on an external quant trading system (`QUANT_ROOT`). They are included as reference implementations.
- `stock-screener` requires a running FastAPI service (part of the quant system) to execute.
- `stock-compare` references external scripts (`market_quote.py`, `run_backtest.py`) from the quant system.

## License

Private repository. Cathay Capital internal use.
