"""
Cathay Capital PE Financial Model — Formula Engine.

Generates all Excel formula strings using the row_map, never hardcoding row numbers.
Each sheet function returns a dict of {row_number: formula_string} for a given column.
The template_builder iterates these dicts and writes them into the workbook.
"""

import sys
import os

_LIB_DIR = os.path.dirname(os.path.abspath(__file__))
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

from row_map import ROWS, SHEETS, row, cell_ref, sheet_cell_ref, data_range
from constants import (COL_HIST_START, COL_HIST_END, COL_FORECAST_START, COL_FORECAST_END,
                       col_letter)
from openpyxl.utils import get_column_letter


# =============================================================================
# 1. Key Assumptions Formulas
# =============================================================================

def assumptions_formulas(col):
    """Formulas for Key Assumptions sheet.
    Most rows are inputs (no formula). Only calculated rows get formulas.
    """
    f = {}
    c = col_letter(col)
    r = ROWS['assumptions']

    # Segment revenues = volume x price x utilization
    f[r['seg_a_revenue']] = f"={c}{r['seg_a_volume']}*{c}{r['seg_a_price']}*{c}{r['seg_a_utilization']}"
    f[r['seg_b_revenue']] = f"={c}{r['seg_b_volume']}*{c}{r['seg_b_price']}*{c}{r['seg_b_utilization']}"
    f[r['seg_c_revenue']] = f"={c}{r['seg_c_volume']}*{c}{r['seg_c_price']}*{c}{r['seg_c_utilization']}"

    # Total revenue = sum of segments
    f[r['total_revenue']] = f"={c}{r['seg_a_revenue']}+{c}{r['seg_b_revenue']}+{c}{r['seg_c_revenue']}"

    # Growth rates (need prior column)
    pc = col_letter(col - 1)
    f[r['total_growth']] = f'=IF({pc}{r["total_revenue"]}=0,"",({c}{r["total_revenue"]}/{pc}{r["total_revenue"]}-1))'
    f[r['seg_a_growth']] = f'=IF({pc}{r["seg_a_revenue"]}=0,"",({c}{r["seg_a_revenue"]}/{pc}{r["seg_a_revenue"]}-1))'
    f[r['seg_b_growth']] = f'=IF({pc}{r["seg_b_revenue"]}=0,"",({c}{r["seg_b_revenue"]}/{pc}{r["seg_b_revenue"]}-1))'
    f[r['seg_c_growth']] = f'=IF({pc}{r["seg_c_revenue"]}=0,"",({c}{r["seg_c_revenue"]}/{pc}{r["seg_c_revenue"]}-1))'

    # Revenue mix percentages
    f[r['seg_a_pct']] = f'=IF({c}{r["total_revenue"]}=0,"",{c}{r["seg_a_revenue"]}/{c}{r["total_revenue"]})'
    f[r['seg_b_pct']] = f'=IF({c}{r["total_revenue"]}=0,"",{c}{r["seg_b_revenue"]}/{c}{r["total_revenue"]})'
    f[r['seg_c_pct']] = f'=IF({c}{r["total_revenue"]}=0,"",{c}{r["seg_c_revenue"]}/{c}{r["total_revenue"]})'

    # Blended COGS % and gross margin
    f[r['blended_cogs_pct']] = (
        f'=IF({c}{r["total_revenue"]}=0,"",'
        f'({c}{r["seg_a_revenue"]}*{c}{r["cogs_pct_a"]}'
        f'+{c}{r["seg_b_revenue"]}*{c}{r["cogs_pct_b"]}'
        f'+{c}{r["seg_c_revenue"]}*{c}{r["cogs_pct_c"]})'
        f'/{c}{r["total_revenue"]})'
    )
    f[r['gross_margin']] = f'=IF({c}{r["blended_cogs_pct"]}="","",1-{c}{r["blended_cogs_pct"]})'

    # Total SG&A %
    f[r['total_sga_pct']] = (
        f"={c}{r['personnel_pct']}+{c}{r['rent_pct']}+{c}{r['marketing_pct']}"
        f"+{c}{r['rd_pct']}+{c}{r['other_opex_pct']}"
    )

    return f


# =============================================================================
# 2. Income Statement Formulas
# =============================================================================

def income_statement_formulas(col):
    """IS formulas — all linked from other sheets via green references."""
    f = {}
    c = col_letter(col)
    r = ROWS['income_statement']

    # Revenue = link from Key Assumptions total_revenue
    ka_sheet = SHEETS['assumptions']['name']
    ka = ROWS['assumptions']
    f[r['revenue']] = f"='{ka_sheet}'!{c}{ka['total_revenue']}"

    # COGS = link from COGS & OpEx
    co_sheet = SHEETS['cogs_opex']['name']
    co = ROWS['cogs_opex']
    f[r['cogs']] = f"='{co_sheet}'!{c}{co['total_cogs']}"

    # Gross Profit = Revenue - COGS
    f[r['gross_profit']] = f"={c}{r['revenue']}-{c}{r['cogs']}"
    f[r['gross_margin']] = f'=IF({c}{r["revenue"]}=0,"",{c}{r["gross_profit"]}/{c}{r["revenue"]})'

    # SG&A = link from COGS & OpEx
    f[r['sga']] = f"='{co_sheet}'!{c}{co['total_sga']}"

    # EBITDA = Gross Profit - SG&A + Other Income
    f[r['ebitda']] = f"={c}{r['gross_profit']}-{c}{r['sga']}+{c}{r['other_income']}"
    f[r['ebitda_margin']] = f'=IF({c}{r["revenue"]}=0,"",{c}{r["ebitda"]}/{c}{r["revenue"]})'

    # D&A = link from COGS & OpEx
    f[r['da']] = f"='{co_sheet}'!{c}{co['total_da']}"

    # EBIT = EBITDA - D&A
    f[r['ebit']] = f"={c}{r['ebitda']}-{c}{r['da']}"
    f[r['ebit_margin']] = f'=IF({c}{r["revenue"]}=0,"",{c}{r["ebit"]}/{c}{r["revenue"]})'

    # Interest = link from Debt & CapEx
    dc_sheet = SHEETS['debt_capex']['name']
    dc = ROWS['debt_capex']
    f[r['interest_expense']] = f"='{dc_sheet}'!{c}{dc['total_interest']}"

    # EBT = EBIT - Interest + Other Financial
    f[r['ebt']] = f"={c}{r['ebit']}-{c}{r['interest_expense']}+{c}{r['other_fin']}"

    # Tax = EBT x tax rate (if EBT > 0, else 0)
    ka_tax = f"'{ka_sheet}'!{c}{ka['tax_rate']}"
    f[r['tax']] = f"=IF({c}{r['ebt']}>0,{c}{r['ebt']}*{ka_tax},0)"
    f[r['effective_tax_rate']] = f'=IF({c}{r["ebt"]}=0,"",{c}{r["tax"]}/{c}{r["ebt"]})'

    # Net Income = EBT - Tax
    f[r['net_income']] = f"={c}{r['ebt']}-{c}{r['tax']}"
    f[r['net_margin']] = f'=IF({c}{r["revenue"]}=0,"",{c}{r["net_income"]}/{c}{r["revenue"]})'

    return f


# =============================================================================
# 3. Balance Sheet Formulas
# =============================================================================

def balance_sheet_formulas(col):
    """BS formulas — derived from IS, WC, Debt, CF."""
    f = {}
    c = col_letter(col)
    pc = col_letter(col - 1)  # prior column
    r = ROWS['balance_sheet']

    # Cash = from CF statement ending cash
    cf_sheet = SHEETS['cash_flow']['name']
    cf = ROWS['cash_flow']
    f[r['cash']] = f"='{cf_sheet}'!{c}{cf['ending_cash']}"

    # AR, Inventory = from Working Capital sheet
    wc_sheet = SHEETS['working_capital']['name']
    wc = ROWS['working_capital']
    f[r['accounts_receivable']] = f"='{wc_sheet}'!{c}{wc['ar_balance']}"
    f[r['inventory']] = f"='{wc_sheet}'!{c}{wc['inventory_balance']}"

    # Total Current Assets
    f[r['total_current_assets']] = (
        f"={c}{r['cash']}+{c}{r['accounts_receivable']}+{c}{r['inventory']}+{c}{r['other_current']}"
    )

    # PP&E = from Debt & CapEx
    dc_sheet = SHEETS['debt_capex']['name']
    dc = ROWS['debt_capex']
    f[r['ppe_net']] = f"='{dc_sheet}'!{c}{dc['ending_ppe']}"

    # Total Non-Current Assets
    f[r['total_noncurrent_assets']] = (
        f"={c}{r['ppe_net']}+{c}{r['intangibles']}+{c}{r['other_noncurrent']}"
    )

    # Total Assets
    f[r['total_assets']] = f"={c}{r['total_current_assets']}+{c}{r['total_noncurrent_assets']}"

    # AP = from Working Capital
    f[r['accounts_payable']] = f"='{wc_sheet}'!{c}{wc['ap_balance']}"

    # Total Current Liabilities
    f[r['total_current_liab']] = (
        f"={c}{r['accounts_payable']}+{c}{r['st_debt']}+{c}{r['other_current_liab']}"
    )

    # LT Debt from Debt & CapEx
    f[r['lt_debt']] = f"='{dc_sheet}'!{c}{dc['total_debt']}"

    f[r['total_noncurrent_liab']] = f"={c}{r['lt_debt']}+{c}{r['other_noncurrent_liab']}"

    f[r['total_liabilities']] = f"={c}{r['total_current_liab']}+{c}{r['total_noncurrent_liab']}"

    # Retained Earnings = prior RE + Net Income - Dividends
    is_sheet = SHEETS['income_statement']['name']
    is_r = ROWS['income_statement']
    ka_sheet = SHEETS['assumptions']['name']
    ka = ROWS['assumptions']
    f[r['retained_earnings']] = (
        f"={pc}{r['retained_earnings']}"
        f"+'{is_sheet}'!{c}{is_r['net_income']}"
        f"-'{is_sheet}'!{c}{is_r['net_income']}*'{ka_sheet}'!{c}{ka['dividend_payout']}"
    )

    f[r['total_equity']] = f"={c}{r['paid_in_capital']}+{c}{r['retained_earnings']}"
    f[r['total_le']] = f"={c}{r['total_liabilities']}+{c}{r['total_equity']}"

    # BS CHECK (must = 0)
    f[r['bs_check']] = f"=ROUND({c}{r['total_assets']}-{c}{r['total_le']},2)"

    return f


# =============================================================================
# 4. Cash Flow Statement Formulas
# =============================================================================

def cash_flow_formulas(col):
    """CF indirect method — linked from IS, WC, Debt."""
    f = {}
    c = col_letter(col)
    pc = col_letter(col - 1)
    r = ROWS['cash_flow']

    # Operating
    is_sheet = SHEETS['income_statement']['name']
    is_r = ROWS['income_statement']
    co_sheet = SHEETS['cogs_opex']['name']
    co = ROWS['cogs_opex']
    wc_sheet = SHEETS['working_capital']['name']
    wc = ROWS['working_capital']

    f[r['net_income']] = f"='{is_sheet}'!{c}{is_r['net_income']}"
    f[r['da']] = f"='{co_sheet}'!{c}{co['total_da']}"

    # Delta WC items from Working Capital sheet
    f[r['delta_ar']] = f"=-'{wc_sheet}'!{c}{wc['delta_ar']}"
    f[r['delta_inventory']] = f"=-'{wc_sheet}'!{c}{wc['delta_inventory']}"
    f[r['delta_ap']] = f"='{wc_sheet}'!{c}{wc['delta_ap']}"
    f[r['total_delta_wc']] = f"={c}{r['delta_ar']}+{c}{r['delta_inventory']}+{c}{r['delta_ap']}+{c}{r['delta_other_wc']}"

    f[r['operating_cf']] = f"={c}{r['net_income']}+{c}{r['da']}+{c}{r['total_delta_wc']}"

    # Investing
    dc_sheet = SHEETS['debt_capex']['name']
    dc = ROWS['debt_capex']
    f[r['capex']] = f"=-'{dc_sheet}'!{c}{dc['total_capex']}"
    f[r['investing_cf']] = f"={c}{r['capex']}+{c}{r['other_investing']}"

    # Financing
    f[r['debt_drawdown']] = f"='{dc_sheet}'!{c}{dc['total_drawdown']}"
    f[r['debt_repayment']] = f"=-'{dc_sheet}'!{c}{dc['total_repayment']}"

    # Dividends from IS
    ka_sheet = SHEETS['assumptions']['name']
    ka = ROWS['assumptions']
    f[r['dividends']] = f"=-'{is_sheet}'!{c}{is_r['net_income']}*'{ka_sheet}'!{c}{ka['dividend_payout']}"

    f[r['financing_cf']] = (
        f"={c}{r['debt_drawdown']}+{c}{r['debt_repayment']}"
        f"+{c}{r['dividends']}+{c}{r['equity_issuance']}"
    )

    # Summary
    f[r['net_cf']] = f"={c}{r['operating_cf']}+{c}{r['investing_cf']}+{c}{r['financing_cf']}"
    f[r['beginning_cash']] = f"={pc}{r['ending_cash']}"
    f[r['ending_cash']] = f"={c}{r['beginning_cash']}+{c}{r['net_cf']}"

    # Cash tie-out check
    bs_sheet = SHEETS['balance_sheet']['name']
    bs = ROWS['balance_sheet']
    f[r['cash_tieout']] = f"=ROUND({c}{r['ending_cash']}-'{bs_sheet}'!{c}{bs['cash']},2)"

    return f


# =============================================================================
# 5. Working Capital Formulas
# =============================================================================

def working_capital_formulas(col):
    """WC schedule — days -> balances -> changes."""
    f = {}
    c = col_letter(col)
    pc = col_letter(col - 1)
    r = ROWS['working_capital']

    # Revenue and COGS references
    ka_sheet = SHEETS['assumptions']['name']
    ka = ROWS['assumptions']
    co_sheet = SHEETS['cogs_opex']['name']
    co = ROWS['cogs_opex']

    revenue_ref = f"'{ka_sheet}'!{c}{ka['total_revenue']}"
    cogs_ref = f"'{co_sheet}'!{c}{co['total_cogs']}"

    # Balances
    f[r['ar_balance']] = f"={revenue_ref}/365*{c}{r['ar_days']}"
    f[r['inventory_balance']] = f"={cogs_ref}/365*{c}{r['inventory_days']}"
    f[r['ap_balance']] = f"={cogs_ref}/365*{c}{r['ap_days']}"
    f[r['net_wc']] = f"={c}{r['ar_balance']}+{c}{r['inventory_balance']}-{c}{r['ap_balance']}"

    # Changes (delta)
    f[r['delta_ar']] = f"={c}{r['ar_balance']}-{pc}{r['ar_balance']}"
    f[r['delta_inventory']] = f"={c}{r['inventory_balance']}-{pc}{r['inventory_balance']}"
    f[r['delta_ap']] = f"={c}{r['ap_balance']}-{pc}{r['ap_balance']}"
    f[r['delta_net_wc']] = f"={c}{r['net_wc']}-{pc}{r['net_wc']}"

    return f


# =============================================================================
# 6. COGS & OpEx Formulas
# =============================================================================

def cogs_opex_formulas(col):
    """COGS by segment + SG&A breakdown + D&A."""
    f = {}
    c = col_letter(col)
    r = ROWS['cogs_opex']
    ka_sheet = SHEETS['assumptions']['name']
    ka = ROWS['assumptions']

    # COGS by segment = segment revenue x COGS %
    f[r['cogs_seg_a']] = f"='{ka_sheet}'!{c}{ka['seg_a_revenue']}*'{ka_sheet}'!{c}{ka['cogs_pct_a']}"
    f[r['cogs_seg_b']] = f"='{ka_sheet}'!{c}{ka['seg_b_revenue']}*'{ka_sheet}'!{c}{ka['cogs_pct_b']}"
    f[r['cogs_seg_c']] = f"='{ka_sheet}'!{c}{ka['seg_c_revenue']}*'{ka_sheet}'!{c}{ka['cogs_pct_c']}"
    f[r['total_cogs']] = f"={c}{r['cogs_seg_a']}+{c}{r['cogs_seg_b']}+{c}{r['cogs_seg_c']}"

    rev_ref = f"'{ka_sheet}'!{c}{ka['total_revenue']}"
    f[r['cogs_pct_rev']] = f'=IF({rev_ref}=0,"",{c}{r["total_cogs"]}/{rev_ref})'

    # Gross Profit
    f[r['gross_profit']] = f"={rev_ref}-{c}{r['total_cogs']}"
    f[r['gross_margin']] = f'=IF({rev_ref}=0,"",{c}{r["gross_profit"]}/{rev_ref})'

    # SG&A = revenue x each SG&A pct
    f[r['personnel']] = f"={rev_ref}*'{ka_sheet}'!{c}{ka['personnel_pct']}"
    f[r['rent']] = f"={rev_ref}*'{ka_sheet}'!{c}{ka['rent_pct']}"
    f[r['marketing']] = f"={rev_ref}*'{ka_sheet}'!{c}{ka['marketing_pct']}"
    f[r['rd']] = f"={rev_ref}*'{ka_sheet}'!{c}{ka['rd_pct']}"
    f[r['other_opex']] = f"={rev_ref}*'{ka_sheet}'!{c}{ka['other_opex_pct']}"
    f[r['total_sga']] = f"={c}{r['personnel']}+{c}{r['rent']}+{c}{r['marketing']}+{c}{r['rd']}+{c}{r['other_opex']}"
    f[r['sga_pct_rev']] = f'=IF({rev_ref}=0,"",{c}{r["total_sga"]}/{rev_ref})'

    # D&A from Debt & CapEx sheet
    dc_sheet = SHEETS['debt_capex']['name']
    dc = ROWS['debt_capex']
    f[r['depreciation']] = f"='{dc_sheet}'!{c}{dc['less_da']}"
    f[r['amortization']] = 0  # placeholder, usually zero for PE deals
    f[r['total_da']] = f"={c}{r['depreciation']}+{c}{r['amortization']}"

    # Total OpEx
    f[r['total_opex']] = f"={c}{r['total_sga']}+{c}{r['total_da']}"

    # EBITDA Bridge
    f[r['ebitda']] = f"={c}{r['gross_profit']}-{c}{r['total_sga']}"
    f[r['ebitda_margin']] = f'=IF({rev_ref}=0,"",{c}{r["ebitda"]}/{rev_ref})'
    f[r['ebit']] = f"={c}{r['ebitda']}-{c}{r['total_da']}"
    f[r['ebit_margin']] = f'=IF({rev_ref}=0,"",{c}{r["ebit"]}/{rev_ref})'

    return f


# =============================================================================
# 7. Debt & CapEx Formulas
# =============================================================================

def debt_capex_formulas(col):
    """Debt schedule + CapEx + PP&E roll-forward."""
    f = {}
    c = col_letter(col)
    pc = col_letter(col - 1)
    r = ROWS['debt_capex']
    ka_sheet = SHEETS['assumptions']['name']
    ka = ROWS['assumptions']

    # Debt Tranche 1
    f[r['d1_beginning']] = f"={pc}{r['d1_ending']}"
    f[r['d1_ending']] = f"={c}{r['d1_beginning']}+{c}{r['d1_drawdown']}-{c}{r['d1_repayment']}"
    f[r['d1_interest']] = f"={c}{r['d1_ending']}*{c}{r['d1_rate']}"

    # Debt Tranche 2
    f[r['d2_beginning']] = f"={pc}{r['d2_ending']}"
    f[r['d2_ending']] = f"={c}{r['d2_beginning']}+{c}{r['d2_drawdown']}-{c}{r['d2_repayment']}"
    f[r['d2_interest']] = f"={c}{r['d2_ending']}*{c}{r['d2_rate']}"

    # Summary
    f[r['total_debt']] = f"={c}{r['d1_ending']}+{c}{r['d2_ending']}"
    f[r['total_interest']] = f"={c}{r['d1_interest']}+{c}{r['d2_interest']}"
    f[r['total_drawdown']] = f"={c}{r['d1_drawdown']}+{c}{r['d2_drawdown']}"
    f[r['total_repayment']] = f"={c}{r['d1_repayment']}+{c}{r['d2_repayment']}"

    # CapEx
    rev_ref = f"'{ka_sheet}'!{c}{ka['total_revenue']}"
    f[r['total_capex']] = f"={c}{r['capex_maintenance']}+{c}{r['capex_growth']}"

    # PP&E Roll-forward
    f[r['beginning_ppe']] = f"={pc}{r['ending_ppe']}"
    f[r['plus_capex']] = f"={c}{r['total_capex']}"
    # D&A = beginning PP&E x DA rate from assumptions
    f[r['less_da']] = f"={c}{r['beginning_ppe']}*'{ka_sheet}'!{c}{ka['da_rate']}"
    f[r['ending_ppe']] = f"={c}{r['beginning_ppe']}+{c}{r['plus_capex']}-{c}{r['less_da']}"

    return f


# =============================================================================
# 8. Returns & Sensitivity Formulas
# =============================================================================

def returns_formulas(col):
    """Return analysis with P/S and P/E exit methods."""
    f = {}
    c = col_letter(col)
    r = ROWS['returns']

    # P/S Exit
    f[r['ps_exit_ev']] = f"={c}{r['ps_exit_revenue']}*{c}{r['ps_exit_multiple']}"
    f[r['ps_exit_equity']] = f"={c}{r['ps_exit_ev']}-{c}{r['ps_exit_net_debt']}"
    f[r['ps_moic']] = f'=IF({c}{r["entry_equity"]}=0,"",{c}{r["ps_exit_equity"]}/{c}{r["entry_equity"]})'

    # P/E Exit
    f[r['pe_exit_equity']] = f"={c}{r['pe_exit_ni']}*{c}{r['pe_exit_multiple']}"
    f[r['pe_moic']] = f'=IF({c}{r["entry_equity"]}=0,"",{c}{r["pe_exit_equity"]}/{c}{r["entry_equity"]})'

    return f


# =============================================================================
# 9. DCF Formulas (with mid-year convention)
# =============================================================================

def dcf_formulas(col, period_num):
    """DCF valuation — WACC + UFCF + Terminal Value.
    period_num: 1-based forecast year index (1 for first forecast year).
    Uses mid-year convention: period = 0.5, 1.5, 2.5, ...
    """
    f = {}
    c = col_letter(col)
    r = ROWS['dcf']

    is_sheet = SHEETS['income_statement']['name']
    is_r = ROWS['income_statement']
    co_sheet = SHEETS['cogs_opex']['name']
    co = ROWS['cogs_opex']
    dc_sheet = SHEETS['debt_capex']['name']
    dc = ROWS['debt_capex']
    wc_sheet = SHEETS['working_capital']['name']
    wc = ROWS['working_capital']

    # WACC (only in first forecast column, shared across)
    wacc_col = col_letter(COL_FORECAST_START)
    if col == COL_FORECAST_START:
        f[r['cost_of_equity']] = f"={c}{r['risk_free']}+{c}{r['beta']}*{c}{r['mrp']}"
        f[r['wacc']] = (
            f"={c}{r['equity_weight']}*{c}{r['cost_of_equity']}"
            f"+{c}{r['debt_weight']}*{c}{r['cost_of_debt']}*(1-{c}{r['tax_rate']})"
        )

    # UFCF
    f[r['ebit']] = f"='{is_sheet}'!{c}{is_r['ebit']}"
    f[r['tax_on_ebit']] = f"={c}{r['ebit']}*{wacc_col}{r['tax_rate']}"
    f[r['nopat']] = f"={c}{r['ebit']}-{c}{r['tax_on_ebit']}"
    f[r['plus_da']] = f"='{co_sheet}'!{c}{co['total_da']}"
    f[r['less_capex']] = f"='{dc_sheet}'!{c}{dc['total_capex']}"
    f[r['less_delta_wc']] = f"='{wc_sheet}'!{c}{wc['delta_net_wc']}"
    f[r['ufcf']] = f"={c}{r['nopat']}+{c}{r['plus_da']}-{c}{r['less_capex']}-{c}{r['less_delta_wc']}"

    # Discount (mid-year convention)
    mid_year = period_num - 0.5
    f[r['discount_period']] = mid_year
    f[r['discount_factor']] = f"=1/(1+${wacc_col}${r['wacc']})^{c}{r['discount_period']}"
    f[r['pv_fcf']] = f"={c}{r['ufcf']}*{c}{r['discount_factor']}"

    return f


# =============================================================================
# 10. Master Dispatch
# =============================================================================

# Map sheet keys to their formula functions
FORMULA_DISPATCH = {
    'assumptions': assumptions_formulas,
    'cogs_opex': cogs_opex_formulas,
    'income_statement': income_statement_formulas,
    'balance_sheet': balance_sheet_formulas,
    'cash_flow': cash_flow_formulas,
    'working_capital': working_capital_formulas,
    'debt_capex': debt_capex_formulas,
    'returns': returns_formulas,
    # dcf uses dcf_formulas(col, period_num) — called separately
    # revenue, comps, dashboard — populated by model_populator
}


def get_formulas(sheet_key, col):
    """Get all formulas for a sheet at a given column.
    Returns dict of {row_number: formula_string}.
    Only returns formulas for forecast columns (H-L).
    Historical columns (D-G) are hardcoded inputs, no formulas.
    """
    if col < COL_FORECAST_START or col > COL_FORECAST_END:
        return {}
    if sheet_key == 'dcf':
        period_num = col - COL_FORECAST_START + 1
        return dcf_formulas(col, period_num)
    fn = FORMULA_DISPATCH.get(sheet_key)
    if fn is None:
        return {}
    return fn(col)


def get_all_formulas():
    """Generate all formulas for all sheets, all forecast columns.
    Returns: dict of {sheet_key: {(row, col): formula_string}}
    """
    all_formulas = {}
    for sheet_key, fn in FORMULA_DISPATCH.items():
        sheet_formulas = {}
        for col in range(COL_FORECAST_START, COL_FORECAST_END + 1):
            col_formulas = fn(col)
            for row_num, formula in col_formulas.items():
                sheet_formulas[(row_num, col)] = formula
        all_formulas[sheet_key] = sheet_formulas

    # DCF needs period_num
    dcf_formulas_all = {}
    for i, col in enumerate(range(COL_FORECAST_START, COL_FORECAST_END + 1)):
        col_formulas = dcf_formulas(col, i + 1)
        for row_num, formula in col_formulas.items():
            dcf_formulas_all[(row_num, col)] = formula
    all_formulas['dcf'] = dcf_formulas_all

    return all_formulas
