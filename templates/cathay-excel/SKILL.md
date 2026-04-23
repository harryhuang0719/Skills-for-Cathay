---
name: cathay-excel-template
description: Cathay Capital PE financial model Excel template. Use when building financial projections, 3-statement models, DCF valuations, or return analyses for PE deal evaluation. Triggers on "financial model", "build a model", "3-statement", "P&L forecast", "revenue projection", "DCF", "returns analysis", "蔚蓝支点模型".
---

# Cathay Excel Financial Model Template

Template: `assets/template.xlsx` (13 sheets, 4 historical + 5 forecast years)

**MUST READ** `docs/specs/2026-04-20-excel-model-rebuild-design.md` for full architecture spec.

**Deal-tested patterns** (CFF scoping, market-size placement, row padding, divider sheets, Return row order, salary inflation) — see `references/ib-deal-patterns.md`. Read this before shipping any model; checklist at bottom of that file.

## Quick Start

```python
import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/cathay-excel-template/lib"))

from template_builder import build_template       # generate fresh template
from model_populator import populate_model         # fill with company data
from data_extractor import extract_from_folder     # parse source materials
from validate_model import validate_model          # 10-point validation
```

### Typical Workflow

```python
# 1. Extract data from source folder
data = extract_from_folder("~/Desktop/340-蔚蓝支点/")

# 2. Review + adjust key assumptions with user
# data['assumptions']['revenue_growth'] = [0.30, 0.25, ...]

# 3. Populate model
populate_model("template.xlsx", data, "蔚蓝支点_Model.xlsx")

# 4. Validate
validate_model("蔚蓝支点_Model.xlsx")
```

## Module Reference

| Module | File | Purpose |
|--------|------|---------|
| `constants` | `lib/constants.py` | Brand colors, column layout, number formats, style presets |
| `row_map` | `lib/row_map.py` | **Declarative row mapping** — all 13 sheets, kills row-offset bugs |
| `formula_engine` | `lib/formula_engine.py` | 617 Excel formulas for 3-statement linkage, DCF, returns |
| `format_engine` | `lib/format_engine.py` | Cathay brand formatting, conditional formats, print areas |
| `template_builder` | `lib/template_builder.py` | Generates clean `template.xlsx` from row_map + formulas |
| `model_populator` | `lib/model_populator.py` | Fills template with data_dict → complete model |
| `data_extractor` | `lib/data_extractor.py` | Parses .xlsx/.docx/.md source materials → data_dict |
| `validate_model` | `lib/validate_model.py` | 10-point validation (BS balance, cash tie-out, formula integrity) |

## Sheet Architecture (13 Sheets)

### Input Layer
| # | Sheet | Purpose |
|---|-------|---------|
| 1 | Cover | Company name, industry, date, analyst, FX |
| 2 | Key Assumptions | **Primary input**: 3 segments (vol×price×util), COGS%, SG&A%, CapEx, WC days, tax, scenario toggle |

### Calculation Layer
| # | Sheet | Purpose |
|---|-------|---------|
| 3 | Revenue Build-up | Bottom-up vs Top-down reconciliation, >10% flag |
| 4 | COGS & OpEx | COGS by segment + SG&A breakdown + D&A |
| 5 | Income Statement | Full P&L, all formula-linked |
| 6 | Balance Sheet | Full BS + auto-check (A = L + E) |
| 7 | Cash Flow Statement | Indirect CF + cash tie-out to BS |
| 8 | Working Capital | AR/AP/Inventory days → ΔWC |
| 9 | Debt & CapEx | Debt tranches + CapEx + PP&E roll-forward |

### Analysis Layer
| # | Sheet | Purpose |
|---|-------|---------|
| 10 | Returns & Sensitivity | P/E + P/S exit → IRR/MOIC + sensitivity tables |
| 11 | DCF Valuation | WACC + UFCF + TV (mid-year convention). Reference only |
| 12 | Comps | 12-company trading comps (EV/Rev, EV/EBITDA, EV/EBIT, P/E) + mean/median |
| 13 | Dashboard | KPI summary + BS/Cash check + returns snapshot |

## Column Layout

```
A: Labels (w=35)   B: Units (w=8)   C: Notes (w=15)
D-G: Historical 2021-2024 (blue font #0070C0, hardcoded inputs)
  G: thick right border (hist/forecast divider)
H-L: Forecast 2025E-2029E (black font, ALL formula-driven)
```

## Key Formula Logic

### 3-Statement Linkage
```
IS: Revenue(from Assumptions) → GP = Rev-COGS → EBITDA = GP-SGA → EBIT = EBITDA-DA → NI = EBT-Tax
BS: Cash(from CF) | AR/Inv(from WC) | PPE(from D&C) | Debt(from D&C) | RE = prior+NI-Div
CF: NI + DA ± ΔWC - CapEx ± Financing → Ending Cash → BS Cash (circular link closed)
```

### Validation Checks (10-point)
1. Sheet structure (13 sheets, correct names)
2. BS Balance (Total Assets = Total L+E)
3. Cash Tie-out (CF ending cash = BS cash)
4. Revenue Bridge (IS revenue = Assumptions total)
5. EBITDA Bridge (formula present)
6. RE Roll-forward (RE = prior + NI - Div)
7. Debt Consistency (BS debt = D&C debt)
8. Formula Presence (anchor cells have formulas)
9. No Blank Anchors (critical cells populated)
10. Cross-sheet Links (green links resolve)

## Formatting Standards

| Element | Format |
|---------|--------|
| Header | `#800000` Cathay red + white bold |
| Sub-header | `#FBE9E8` light pink + dark red bold |
| Historical input | Blue font `#0070C0` |
| Formula cells | Black font |
| Cross-sheet links | Green font `#00B050` |
| Alternating rows | `#F2F2F2` fill |
| Numbers | `#,##0` |
| Percentages | `0.0%` |
| Negatives | Red parentheses `(#,##0)` |
| BS/Cash check | Green = balanced, Red = error |
| Hist/Forecast divider | Thick right border on column G |

## How to Use

### Option A: Full automation (from source folder)
```
"帮蔚蓝支点做财务模型，文件在 ~/Desktop/340-蔚蓝支点/"
→ data_extractor scans folder
→ present key assumptions for user review
→ model_populator generates complete model
→ validate_model checks everything
```

### Option B: Manual assumptions
```
"做一个3-segment的财务模型，收入1.2亿，毛利35%，预期年增长30%"
→ build data_dict from conversation
→ populate_model fills template
→ validate_model checks
```

### Option C: Template only
```
"给我一个空白的Cathay财务模型模板"
→ copy assets/template.xlsx
→ user fills blue input cells manually
```

## Regenerating Template

```python
from template_builder import build_template
build_template()  # generates assets/template.xlsx
```

## data_dict Format

```python
{
    'company_name': '蔚蓝支点',
    'industry': '核能/SMR',
    'segments': ['设备销售', '运维服务', '技术授权'],
    'historical': {
        2021: {'revenue': 50, 'cogs': 30, 'sga': 8, 'da': 3,
               'interest': 1, 'tax': 2, 'cash': 20, 'ar': 15,
               'inventory': 10, 'ppe': 50, 'ap': 12, 'debt': 30,
               'equity': 40, 'retained_earnings': 15, 'capex': 8},
        # ... 2022, 2023, 2024
    },
    'assumptions': {
        'revenue_growth': [0.30, 0.25, 0.20, 0.15, 0.12],
        'gross_margin_target': [0.42, 0.44, 0.45, 0.46, 0.47],
        'sga_pct': 0.12,
        'capex_pct': 0.08,
        'tax_rate': 0.25,
        'ar_days': 60, 'ap_days': 45, 'inventory_days': 30,
        'da_rate': 0.10,
        'dividend_payout': 0.0,
        'interest_rate': 0.05,
    },
}
```
