# IB Model Deal-Tested Patterns

Patterns captured from live deal modeling work where the generic template
was correct but specific modeling choices were wrong. Each entry names a
concrete mistake and the fix.

Source: 深海智人 / Project Sealien (2026-04). Compare sponsor build
(Cathay_v4.xlsx) vs user's final version (Sealien FM_260423.xlsx).

---

## 1. CFF Equity Raise = Total Round, Not Sponsor's Slice

**Mistake**: Cash inflow from a primary round modeled as the sponsor's check
size only.

**Fix**: The WHOLE round (all investors) flows into the company's cash via
CFF. The sponsor's portion is used only for ownership / dilution math.

```
Terms block (Return sheet, rows 14-16):
  Pre-money    3,000 M
  Total raise    400 M   ← this goes to CFF
  Cathay         200 M   ← this drives Cathay ownership only

Financials CFS "+ 股本/资本公积 新增" formula:
  =Return!C15 * 1000      ← total raise × 1000 (M to 千元)
  NOT: =Return!C16 * 1000  ← Cathay slice only — WRONG
```

Applies any time the sponsor is a minority investor in a round.

## 2. Market Size Rows at Top of P&L, Not Below NI

**Mistake**: Market share / TAM / SAM placed as KPI rows after Net Income.

**Fix**: Put 4 market rows immediately above Revenue (as the P&L opens):

```
r6:  全球市场规模 (亿)       =Market!F11   ← link to Market sheet
r7:  中国市场规模 (亿)       =Market!F22
r8:  全球市占率 %           =Revenue/Global
r9:  中国市占率 %           =Revenue/China
r10: (blank)
r11: 营业收入 (Revenue)
```

Reader opens Financials → first sees "size of prize" → then reads how
company penetrates. This is IB convention — market context frames the P&L.

## 3. No Row-Number Padding in Financials

**Mistake**: Force anchor rows with `while r < N: r += 1` padding so outside
sheets can reference `Financials!K36` as a stable NI row.

**Fix**: Use captured row variables everywhere. Let rows flow naturally.
External sheets reference via captured refs, not hardcoded numbers.

```python
# Bad: forces padding and brittle coupling
r = current_row
while r < 36:     # pad to hit "NI at row 36"
    r += 1
r_ni = r

# Good: let rows flow; capture the ref; use it downstream
r_ni = r
# ... later in another sheet:
f'=Financials!{col}{wb._fin_rows["ni"]}'
```

Compacts Financials ~35% (observed 162 → 105 rows in Sealien) and
eliminates silent bugs when adding/removing rows. The only legitimate
reason to pad is if external consumers (not your own code) will hardcode
row numbers — which they shouldn't in a single-workbook model.

## 4. Divider Sheets `>>`

**Pattern**: Insert empty sheets named `>>` (one blank cell) between logical
groups to separate inputs / calculations / appendices:

```
  Return            ← output / summary
  >>
  Assumptions       ← inputs
  Revenue Build
  Financials
  Accounts
  >>
  Market TAM        ← appendix / reference
  Comps
```

openpyxl:
```python
ws = wb.create_sheet('>>')
ws.sheet_properties.tabColor = '808080'  # optional grey tab
```

Users expect this in banker-grade models. Reading flow: "first see the
answer, then the assumptions, then the appendices."

## 5. Return Block Row Order

**Correct order** (NI and Hold grouped as "IRR inputs"; MOIC / IRR at bottom):

```
2030 NI (RMB M)
Hold Period (years)       ← immediately after NI (both feed IRR)
Exit PE (x)
Exit Equity Value (NI × PE)
Cathay Final Ownership %  ← after IPO / full dilution chain
Cathay Exit Equity
Cathay Total Invested
(blank)
MOIC                      ← separated from inputs by blank
IRR
```

Rationale: Exit valuation flows through top-down; MOIC/IRR are the two KPIs
and should be visually isolated. Hold period drives IRR math; putting it
next to NI keeps the time-weighted-return inputs together.

## 6. Title Minimalism

**Pattern**:
```
{Company} — 财务模型 ({YYYY-MM})
```

Avoid organization prefixes on the model title ("Cathay Capital IC 财务模型"
etc.). Org branding lives in slide headers and filenames, not in cell A1
of the Cover sheet. Short titles read better and scale to multiple
projects without looking like boilerplate.

## 7. Salary Inflation in Long-Horizon Models

**Mistake**: Hold 人均工资 (wage per head) flat for 5-7 year forecast.

**Fix**: Grow by 3-5% p.a. Flat salaries overstate operating leverage and
understate OpEx scaling with headcount.

Example (Sealien):
- 管理 人均工资:  300 (2026) → 365 (2030)   ≈ 5% CAGR
- 研发 人均工资:  400 (2026) → 486 (2030)   ≈ 5% CAGR
- 销售 人均工资:  350 (flat — revenue-driving role, kept as anchor)

Not all roles need inflation; sales comp often flat if variable-comp-heavy.
Use judgment per function.

---

## Quick-check Checklist Before Shipping

- [ ] CFF primary-round inflow references TOTAL raise, not sponsor slice
- [ ] Market size / share visible above Revenue in P&L (not only at bottom)
- [ ] No unnecessary row padding (`while r < N: r += 1`) in any sheet
- [ ] Divider `>>` sheets between input / output / appendix groups
- [ ] Return table: NI → Hold → PE → ExitEq → Ownership → Cathay → Invested → MOIC → IRR
- [ ] Title has no organization prefix
- [ ] 人均工资 series grows 3-5% p.a. for non-sales roles in 5+ year forecasts
