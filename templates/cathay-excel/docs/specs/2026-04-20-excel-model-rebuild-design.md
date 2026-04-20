# Cathay Excel Financial Model — Rebuild Design Spec

**Date:** 2026-04-20
**Status:** Approved
**Scope:** Growth Equity + Mid-stage PE (no LBO)

## Problem

Current template has critical issues found by Codex audit:
- 12 sheets shipped (SKILL.md claims 20), formula row offsets corrupt core calculations
- 3 conflicting builder scripts, no canonical generator
- Revenue, COGS, IS formulas reference wrong rows
- 3-statement linkage broken (D&A, interest, cash, PP&E all zero placeholders)
- Return Analysis has zero formulas
- validate_model() exists only in docs, not in code
- No lib/ structure, no tests

## Design

### Input/Output Contract

**Input:** User provides:
1. Local folder path containing source materials (.xlsx, .docx, .md)
2. Key assumptions via dialogue:
   - 3 segment definitions + historical revenue split
   - Expected revenue growth / margin trajectory
   - Core asset/debt positions
   - CapEx plan

**Output:**
- `CompanyName_Model.xlsx` — fully linked 13-sheet model with formulas
- Validation report (BS balance, cash tie-out, formula integrity)

### File Architecture

```
cathay-excel-template/
├── SKILL.md                    ← updated skill definition
├── assets/
│   └── template.xlsx           ← regenerated clean template
├── lib/
│   ├── __init__.py
│   ├── constants.py            ← brand colors, number formats, sheet names
│   ├── row_map.py              ← declarative row mapping (kills row-offset bugs)
│   ├── formula_engine.py       ← all Excel formula generation (3-statement core)
│   ├── format_engine.py        ← Cathay brand formatting, conditional formats
│   ├── template_builder.py     ← generates template.xlsx from row_map + formulas
│   ├── model_populator.py      ← fills template with data_dict → complete model
│   ├── data_extractor.py       ← extracts financials from .xlsx/.docx/.md files
│   └── validate_model.py       ← BS balance, cash tie-out, formula checks, regression
├── tests/
│   └── test_template.py        ← opens generated xlsx, asserts formulas in anchor cells
├── docs/
│   └── specs/
│       └── this file
└── references/
    └── formula-reference.md    ← formula logic documentation for auditors
```

### Sheet Architecture (13 Sheets)

#### Input Layer (2 sheets)
| # | Sheet | Rows | Content |
|---|-------|------|---------|
| 1 | Cover | ~15 | Company, industry, date, analyst, FX rate |
| 2 | Key Assumptions | ~100 | **Primary input sheet.** 3 segments × (volume, price, utilization), COGS % by segment, SG&A breakdown, CapEx, WC days, tax rate, scenario toggle (Base/Upside/Downside) |

#### Calculation Layer (7 sheets)
| # | Sheet | Rows | Content |
|---|-------|------|---------|
| 3 | Revenue Build-up | ~40 | Bottom-up (Vol×Price×Util per seg) vs Top-down (TAM×penetration×share), reconciliation with >10% flag |
| 4 | COGS & OpEx | ~50 | Materials + labor + overhead + D&A by segment; SG&A breakdown (personnel, rent, marketing, other) |
| 5 | Income Statement | ~35 | Revenue → Gross Profit → EBITDA → EBIT → EBT → Net Income. All formula-linked. |
| 6 | Balance Sheet | ~45 | Current/Non-current assets + liabilities + equity. Auto-check row (A = L + E). Retained earnings = prior + NI - dividends |
| 7 | Cash Flow Statement | ~35 | Indirect method: NI + D&A ± ΔWC - CapEx ± financing. Cash tie-out to BS |
| 8 | Working Capital | ~30 | AR/AP/Inventory: historical days (blue input), forecast = avg of last 3 historical. ΔWC auto-calculated |
| 9 | Debt & CapEx | ~50 | Debt tranches (up to 3) with amortization + interest. CapEx by category + depreciation schedule |

#### Analysis Layer (4 sheets)
| # | Sheet | Rows | Content |
|---|-------|------|---------|
| 10 | Returns & Sensitivity | ~55 | **P/E exit + P/S exit → IRR/MOIC** (IPO & M&A scenarios). Sensitivity tables: IRR vs entry×exit multiples; Revenue CAGR vs exit year |
| 11 | DCF Valuation | ~40 | WACC (CAPM) + UFCF + Terminal Value (Gordon/Exit Multiple toggle). **Reference only**, not primary valuation |
| 12 | Comps | ~45 | 12-company table: EV/Rev, EV/EBITDA, EV/EBIT, P/E. Mean/Median row |
| 13 | Dashboard | ~30 | KPI summary, P&L snapshot, returns highlight, BS/Cash check status |

### Column Layout (unchanged)

```
A: Labels (w=35)   B: Units (w=8)   C: Notes (w=15)
D-G: Historical 2021-2024 (blue font, hardcoded inputs)
  G: thick right border (hist/forecast divider)
H-L: Forecast 2025E-2029E (black font, all formulas)
```

### row_map.py — Declarative Row Mapping

The #1 architectural improvement. Every sheet has a dict mapping semantic names to row numbers. Formula generation references these names, never raw integers.

```python
SHEETS = {
    'cover': {'name': 'Cover', 'index': 0},
    'assumptions': {'name': 'Key Assumptions', 'index': 1},
    'revenue': {'name': 'Revenue Build-up', 'index': 2},
    ...
}

ROWS = {
    'income_statement': {
        'header': 1,
        'year_row': 3,
        'revenue': 5,
        'cogs': 8,
        'gross_profit': 10,       # = revenue - cogs
        'gross_margin': 11,       # = gross_profit / revenue
        'sga_total': 14,
        'other_income': 16,
        'ebitda': 18,             # = gross_profit - sga_total + other_income
        'ebitda_margin': 19,
        'da': 21,
        'ebit': 22,               # = ebitda - da
        'interest_expense': 24,
        'ebt': 26,                # = ebit - interest
        'tax': 27,
        'tax_rate': 28,
        'net_income': 30,         # = ebt - tax
        'net_margin': 31,
    },
    'balance_sheet': {
        'header': 1,
        'year_row': 3,
        # Current Assets
        'cash': 5,
        'accounts_receivable': 6,
        'inventory': 7,
        'other_current': 8,
        'total_current_assets': 10,
        # Non-Current Assets
        'ppe_net': 12,
        'intangibles': 13,
        'other_noncurrent': 14,
        'total_noncurrent_assets': 16,
        'total_assets': 18,
        # Current Liabilities
        'accounts_payable': 21,
        'st_debt': 22,
        'other_current_liab': 23,
        'total_current_liab': 25,
        # Non-Current Liabilities
        'lt_debt': 27,
        'other_noncurrent_liab': 28,
        'total_noncurrent_liab': 30,
        'total_liabilities': 32,
        # Equity
        'paid_in_capital': 34,
        'retained_earnings': 35,  # = prior RE + NI - dividends
        'total_equity': 37,
        'total_le': 39,           # = total_liabilities + total_equity
        'bs_check': 41,           # = total_assets - total_le (should be 0)
    },
    ...
}
```

### formula_engine.py — Three-Statement Linkage

Core formulas generated from row_map, never hardcoded:

```
IS:
  Gross Profit = Revenue - COGS
  EBITDA = Gross Profit - SG&A + Other Income
  EBIT = EBITDA - D&A
  EBT = EBIT - Interest
  Net Income = EBT - Tax

BS:
  Total Assets = Current Assets + Non-Current Assets
  Total L+E = Total Liabilities + Total Equity
  Retained Earnings = Prior RE + Net Income (from IS)
  Cash = Prior Cash + Net CF (from CF)
  PP&E = Prior PP&E + CapEx - D&A (from Debt & CapEx)
  BS Check = Total Assets - Total L+E (must = 0)

CF (indirect):
  Operating CF = Net Income + D&A ± ΔWC
  Investing CF = -CapEx
  Financing CF = Debt drawdown - Debt repayment - Dividends
  Net CF = Operating + Investing + Financing
  Ending Cash = Beginning Cash + Net CF
  Cash Tie-out = Ending Cash - BS Cash (must = 0)

Returns:
  Exit Equity (P/E) = Net Income × Exit P/E Multiple
  Exit Equity (P/S) = Revenue × Exit P/S Multiple
  IRR = IRR({-entry_equity, 0, 0, ..., exit_equity})
  MOIC = Exit Equity / Entry Equity
```

### Formatting Standards (Cathay brand)

| Element | Spec |
|---------|------|
| Header | `#800000` bg + white bold 11pt |
| Sub-header | `#FBE9E8` light pink bg + dark red bold |
| Historical input | Blue font `#0070C0` |
| Forecast formula | Black font |
| Cross-sheet link | Green font `#00B050` |
| Alternating rows | `#F2F2F2` fill on even rows |
| Numbers | `#,##0` with thousands separator |
| Percentages | `0.0%` |
| Negatives | Red `(#,##0)` parentheses |
| BS/Cash check | Conditional: 0 → green "PASS", ≠0 → red value |
| Hist/Forecast divider | Thick right border on column G |
| Print area | Set on every sheet |

### validate_model.py — Real Validation

Checks beyond BS/cash:
1. **BS Balance** — Total Assets = Total L+E across all years
2. **Cash Tie-out** — CF ending cash = BS cash across all years
3. **Revenue Bridge** — IS revenue = Revenue Build-up total
4. **EBITDA Bridge** — IS EBITDA = Revenue - COGS - SG&A
5. **RE Roll-forward** — Retained Earnings(t) = RE(t-1) + NI(t)
6. **Debt Paydown** — BS debt = Debt Schedule ending balance
7. **DCF Inputs** — UFCF cells not blank/zero
8. **Return Mechanics** — IRR/MOIC cells have formulas, not hardcoded
9. **No Blank Green Cells** — cross-sheet links actually resolve
10. **Formula Presence** — anchor cells in forecast columns contain formulas, not values

### data_extractor.py — Source Material Parsing

Scans a folder and extracts:
- `.xlsx` → openpyxl: find sheets with "P&L", "IS", "BS", financial keywords → extract historical data
- `.docx` → python-docx: extract tables with financial data, key metrics from body text
- `.md` → regex: extract numbers with context (revenue, margin, CAPEX, headcount)

Returns a standardized `data_dict`:
```python
{
    'company_name': '蔚蓝支点',
    'industry': '核能/SMR',
    'segments': ['设备销售', '运维服务', '技术授权'],
    'historical': {
        2021: {'revenue': 50_000_000, 'cogs': 30_000_000, ...},
        2022: {...},
        2023: {...},
        2024: {...},
    },
    'assumptions': {
        'revenue_growth': [0.30, 0.25, 0.20, 0.15, 0.12],  # per forecast year
        'gross_margin': [0.42, 0.44, 0.45, 0.46, 0.47],
        'sga_pct': 0.12,
        'capex_pct': 0.08,
        'tax_rate': 0.25,
        'ar_days': 60,
        'ap_days': 45,
        'inventory_days': 30,
    },
}
```

### Workflow

```
User: "帮蔚蓝支点做财务模型，文件在 ~/Desktop/340-蔚蓝支点/"

Step 1: data_extractor scans folder → extracts historical financials
Step 2: Present findings + ask for key assumptions:
  - "找到2021-2024历史收入/成本数据，请确认3个segment定义"
  - "预期收入增长率？毛利率变化趋势？"
  - "CapEx计划？融资假设？"
Step 3: User confirms/adjusts assumptions
Step 4: model_populator generates complete model
Step 5: validate_model runs all 10 checks
Step 6: Output .xlsx + validation report
```

## Implementation Phases

| Phase | Scope | Files |
|-------|-------|-------|
| 1 | constants + row_map + formula_engine | Foundation: all formulas correct by construction |
| 2 | format_engine + template_builder | Generate clean template.xlsx |
| 3 | validate_model + test_template | Regression tests: assert formulas in anchor cells |
| 4 | model_populator | Fill template with data_dict |
| 5 | data_extractor | Parse .xlsx/.docx/.md source materials |
| 6 | SKILL.md rewrite | Updated documentation |
