#!/usr/bin/env python3
"""Build the ASM International Valuation Model workbook."""

import sys
import os
import datetime
import copy as _copy

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from scripts.formatting import (
    BLUE_FONT,
    BLACK_FONT,
    BLACK_BOLD,
    GREEN_FONT,
    YELLOW_FILL,
    FMT_CURRENCY,
    FMT_PERCENT,
    FMT_YEAR,
    FMT_RATIO,
    FMT_MULTIPLE,
    style_input_cell,
    style_formula_cell,
    style_crossref_cell,
    style_assumption_cell,
    style_section_header,
    style_total_row,
    style_double_line_row,
    write_year_headers,
    write_label,
    set_column_widths,
    freeze_panes,
)
from openpyxl.utils import get_column_letter
from scripts.data_maps import SOURCE_FILE, COMPANY_INFO, YEARS

OUTPUT_FILE = "output/ASM_Valuation_Model.xlsx"

# ---------------------------------------------------------------------------
# Input tab mappings: (source sheet name, target sheet name)
# ---------------------------------------------------------------------------
INPUT_TAB_MAP = [
    ("ASM", "Input ASM"),
    ("ASM (as reported)", "Input ASM (Reported)"),
    ("AIXTRON", "Input AIXTRON"),
    ("Applied Materials", "Input AMAT"),
    ("Lam Research", "Input LRCX"),
]


# ---------------------------------------------------------------------------
# Cover tab
# ---------------------------------------------------------------------------
def build_cover(wb):
    """Build the Cover (title page) tab."""
    # Use the default sheet created with the workbook
    ws = wb.active
    ws.title = "Cover"

    # Column widths
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 40

    # ---- Title block ----
    row = 3
    cell = ws.cell(row=row, column=2, value="ASM International NV")
    cell.font = Font(size=22, bold=True, color="000000")
    cell.alignment = Alignment(horizontal="left")

    row = 4
    cell = ws.cell(row=row, column=2, value="Equity Valuation Report")
    cell.font = Font(size=16, bold=False, color="333333")
    cell.alignment = Alignment(horizontal="left")

    row = 6
    cell = ws.cell(row=row, column=2, value="AC444 Valuation and Security Analysis")
    cell.font = Font(size=12, bold=False, color="000000")

    row = 7
    cell = ws.cell(row=row, column=2, value="London School of Economics")
    cell.font = Font(size=12, bold=False, color="000000")

    # ---- Key information ----
    row = 9
    info_items = [
        ("Ticker:", "ENXTAM:ASM"),
        ("Currency:", "EUR (Thousands)"),
        ("Data Period:", "FY 2019 - FY 2024"),
        ("Forecast Period:", "FY 2025E - FY 2030E"),
        ("Peers:", "AIXTRON SE, Applied Materials, Lam Research"),
    ]

    label_font = Font(size=11, bold=True, color="000000")
    value_font = Font(size=11, bold=False, color="000000")

    for label, value in info_items:
        ws.cell(row=row, column=2, value=label).font = label_font
        ws.cell(row=row, column=3, value=value).font = value_font
        row += 1

    # ---- Table of Contents ----
    row += 1  # blank row
    cell = ws.cell(row=row, column=2, value="Table of Contents")
    cell.font = Font(size=14, bold=True, color="000000")
    cell.border = Border(bottom=Side(style="medium"))
    ws.cell(row=row, column=3).border = Border(bottom=Side(style="medium"))
    row += 1

    toc_sections = [
        ("1.", "Cover", "Title page and overview"),
        ("2.", "Input ASM", "Raw financial data - ASM International"),
        ("3.", "Input ASM (Reported)", "As-reported financial data - ASM International"),
        ("4.", "Input AIXTRON", "Raw financial data - AIXTRON SE"),
        ("5.", "Input AMAT", "Raw financial data - Applied Materials"),
        ("6.", "Input LRCX", "Raw financial data - Lam Research"),
        ("7.", "Standardized ASM", "Standardized financial statements - ASM"),
        ("8.", "Standardized AIXTRON", "Standardized financial statements - AIXTRON"),
        ("9.", "Standardized AMAT", "Standardized financial statements - Applied Materials"),
        ("10.", "Standardized LRCX", "Standardized financial statements - Lam Research"),
        ("11.", "Adjustments", "Adjustments and adjusted financial statements"),
        ("12.", "Ratio Analysis", "Key financial ratios and analysis"),
        ("13.", "Forecasts", "Revenue, margin, and balance sheet forecasts"),
        ("14.", "Valuation", "WACC, DCF, and PVAOI valuation"),
        ("15.", "Sensitivity", "Sensitivity analysis tables"),
        ("16.", "Comps", "Comparable company analysis"),
        ("17.", "Football Field", "Valuation range summary"),
    ]

    num_font = Font(size=10, bold=True, color="000000")
    section_font = Font(size=10, bold=False, color="000000")
    desc_font = Font(size=10, bold=False, color="666666")

    for num, section, description in toc_sections:
        ws.cell(row=row, column=2, value=f"{num}  {section}").font = num_font
        ws.cell(row=row, column=3, value=description).font = desc_font
        row += 1

    # ---- Date stamp ----
    row += 1
    ws.cell(
        row=row,
        column=2,
        value=f"Generated: {datetime.date.today().strftime('%d %B %Y')}",
    ).font = Font(size=9, italic=True, color="888888")

    # Freeze panes - nothing to freeze on Cover, but set print area nicely
    ws.sheet_properties.tabColor = "1F4E79"

    return ws


# ---------------------------------------------------------------------------
# Copy source sheet helper
# ---------------------------------------------------------------------------
def copy_source_sheet(wb, source_wb, source_sheet_name, target_sheet_name):
    """Copy all non-empty rows from a source sheet into a new target sheet.

    - Column A (labels) = black font
    - Data columns (B onwards) with numeric values = blue font (hard-coded inputs)
    - String values = black font
    - Date values = kept as-is
    - Column A width = 45, data columns = 14
    - Freeze panes at B2
    """
    src_ws = source_wb[source_sheet_name]
    tgt_ws = wb.create_sheet(title=target_sheet_name)

    max_col = src_ws.max_column or 1
    max_row = src_ws.max_row or 1

    for row_idx in range(1, max_row + 1):
        # Check if this row has any non-empty cell
        row_has_data = False
        for col_idx in range(1, max_col + 1):
            val = src_ws.cell(row=row_idx, column=col_idx).value
            if val is not None:
                row_has_data = True
                break

        if not row_has_data:
            continue

        # Copy all cells in this row
        for col_idx in range(1, max_col + 1):
            src_cell = src_ws.cell(row=row_idx, column=col_idx)
            tgt_cell = tgt_ws.cell(row=row_idx, column=col_idx)
            val = src_cell.value
            tgt_cell.value = val

            if val is None:
                continue

            # Column A = labels -> black font
            if col_idx == 1:
                tgt_cell.font = BLACK_FONT
            else:
                # Data columns: numeric -> blue, string -> black, date -> keep
                if isinstance(val, (int, float)):
                    tgt_cell.font = BLUE_FONT
                elif isinstance(val, datetime.datetime):
                    # Keep date as-is (no special font override)
                    pass
                else:
                    tgt_cell.font = BLACK_FONT

    # Set column widths: A=45, data columns=14
    widths = {1: 45}
    for col_idx in range(2, max_col + 1):
        widths[col_idx] = 14
    set_column_widths(tgt_ws, widths=widths)

    # Freeze panes at B2 (row=2, col=2)
    freeze_panes(tgt_ws, row=2, col=2)

    return tgt_ws


# ---------------------------------------------------------------------------
# Build all input tabs
# ---------------------------------------------------------------------------
def build_input_tabs(wb, source_wb):
    """Build all Input tabs by copying data from source sheets."""
    for source_sheet_name, target_sheet_name in INPUT_TAB_MAP:
        print(f"  Copying '{source_sheet_name}' -> '{target_sheet_name}'...")
        copy_source_sheet(wb, source_wb, source_sheet_name, target_sheet_name)


# ---------------------------------------------------------------------------
# Standardized ASM tab
# ---------------------------------------------------------------------------
def build_standardized_tab(wb, company_key, input_sheet_name, reported_sheet_name, output_sheet_name):
    """Build a standardized financial statements tab with reformulated IS, BS, and CF.

    Parameters
    ----------
    wb : openpyxl.Workbook
        The output workbook.
    company_key : str
        Key into COMPANY_INFO (e.g. 'ASM', 'AIXA', 'AMAT', 'LRCX').
    input_sheet_name : str
        Quoted sheet name for formula references (e.g. "'Input ASM'").
    reported_sheet_name : str or None
        Quoted sheet name for as-reported data (e.g. "'Input ASM (Reported)'").
        None for peer companies that have no reported sheet.
    output_sheet_name : str
        Name for the output tab (e.g. 'Std ASM', 'Std AIXTRON').

    The reformulated balance sheet separates operating, investment, and
    financing items.  The accounting identity used is:

        Business Assets (= NOA + Investment Assets)
        = Net Debt + Equity  (= Invested Capital)

    All data cells are Excel formulas referencing the Input tabs.
    """
    ws = wb.create_sheet(title=output_sheet_name)
    MAX_COL = 7  # columns A-G

    # Shortcuts to data maps
    info = COMPANY_INFO[company_key]
    is_map = info['is_map']
    cf_map = info['cf_map']
    bs_map = info['bs_map']

    input_sheet = input_sheet_name

    # --- Column widths ---
    widths = {1: 45}
    for c in range(2, 8):
        widths[c] = 14
    set_column_widths(ws, widths=widths)

    # --- Helper: write a cross-ref formula across years (with ISNUMBER guard) ---
    def write_crossref_row(row, sheet, src_row, fmt=FMT_CURRENCY, guard=True):
        """Write cross-sheet reference formulas across cols B-G."""
        for ci in range(2, 8):
            cl = get_column_letter(ci)
            cell = ws.cell(row=row, column=ci)
            if guard:
                cell.value = f"=IF(ISNUMBER({sheet}!{cl}{src_row}),{sheet}!{cl}{src_row},0)"
            else:
                cell.value = f"={sheet}!{cl}{src_row}"
            style_crossref_cell(cell, fmt)

    def write_zero_row(row, fmt=FMT_CURRENCY):
        """Write 0 across cols B-G for missing map keys."""
        for ci in range(2, 8):
            cell = ws.cell(row=row, column=ci, value=0)
            style_formula_cell(cell, fmt)

    def write_crossref_or_zero(row, sheet, map_dict, key, fmt=FMT_CURRENCY, guard=True):
        """Write crossref if key exists in map, else write 0."""
        if key in map_dict:
            write_crossref_row(row, sheet, map_dict[key], fmt=fmt, guard=guard)
        else:
            write_zero_row(row, fmt=fmt)

    def write_formula_row(row, formulas_by_col, fmt=FMT_CURRENCY, bold=False):
        """Write formula strings across cols B-G. formulas_by_col is a callable(col_letter)->str."""
        for ci in range(2, 8):
            cl = get_column_letter(ci)
            cell = ws.cell(row=row, column=ci)
            cell.value = formulas_by_col(cl)
            style_formula_cell(cell, fmt)
            if bold:
                cell.font = BLACK_BOLD

    # =========================================================================
    # Row 1-4: Title, subtitle, year headers
    # =========================================================================
    r = 1
    cell = ws.cell(row=r, column=1, value=f"Standardized Financial Statements - {info['name']}")
    cell.font = Font(size=14, bold=True, color="000000")

    r = 2
    ws.cell(row=r, column=1, value="(EUR 000)").font = Font(size=10, italic=True, color="666666")

    r = 3
    write_year_headers(ws, r, 2, YEARS)

    r = 4  # blank

    # =========================================================================
    # Section 1: Assumptions (rows 5-7)
    # =========================================================================
    r = 5
    write_label(ws, r, 1, "Operating cash %")
    for ci in range(2, 8):
        cell = ws.cell(row=r, column=ci, value=0.02)
        style_assumption_cell(cell, FMT_PERCENT)

    r = 6
    write_label(ws, r, 1, "Tax rate (effective)")
    write_crossref_or_zero(r, input_sheet, is_map, 'effective_tax_rate', fmt=FMT_PERCENT, guard=True)

    r = 7  # blank

    # =========================================================================
    # Section 2: INCOME STATEMENT (row 8 onwards)
    # =========================================================================
    r = 8
    style_section_header(ws, r, MAX_COL, "INCOME STATEMENT")

    # Row 9: Revenue
    r = 9
    write_label(ws, r, 1, "Revenue")
    write_crossref_row(r, input_sheet, is_map['revenue'])

    # Row 10: Cost of Sales
    r = 10
    write_label(ws, r, 1, "Cost of Sales")
    write_crossref_row(r, input_sheet, is_map['cogs'])

    # Row 11: Gross Profit = Revenue - COGS
    r = 11
    write_label(ws, r, 1, "Gross Profit", bold=True)
    write_formula_row(r, lambda cl: f"={cl}9-{cl}10", bold=True)

    # Row 12: SG&A
    r = 12
    write_label(ws, r, 1, "SG&A", indent=2)
    write_crossref_row(r, input_sheet, is_map['sga'])

    # Row 13: R&D
    r = 13
    write_label(ws, r, 1, "R&D", indent=2)
    write_crossref_row(r, input_sheet, is_map['rd'])

    # Row 14: D&A (from CF supplemental -- memo line, already embedded in COGS)
    r = 14
    write_label(ws, r, 1, "Depreciation & Amortization", indent=2)
    write_crossref_row(r, input_sheet, cf_map['da_total'])

    # Row 15: Other Operating Expense
    r = 15
    write_label(ws, r, 1, "Other Operating Expense", indent=2)
    write_crossref_or_zero(r, input_sheet, is_map, 'other_operating_exp')

    # Row 16: Total Operating Expense = COGS + SGA + RD + DA + Other
    r = 16
    write_label(ws, r, 1, "Total Operating Expense", bold=True)
    write_formula_row(r, lambda cl: f"={cl}10+{cl}12+{cl}13+{cl}14+{cl}15", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 17: Recurring Operating Profit = Revenue - Total OpEx
    r = 17
    write_label(ws, r, 1, "Recurring Operating Profit", bold=True)
    write_formula_row(r, lambda cl: f"={cl}9-{cl}16", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 18: blank
    r = 18

    # Row 19: Non-recurring items header
    r = 19
    write_label(ws, r, 1, "Non-recurring items:", bold=True)

    # Row 20: Gain/Loss on Sale of Investments
    r = 20
    write_label(ws, r, 1, "Gain/Loss on Sale of Investments", indent=2)
    write_crossref_or_zero(r, input_sheet, is_map, 'gain_loss_sale_invest')

    # Row 21: Gain/Loss on Sale of Assets
    r = 21
    write_label(ws, r, 1, "Gain/Loss on Sale of Assets", indent=2)
    write_crossref_or_zero(r, input_sheet, is_map, 'gain_loss_sale_assets')

    # Row 22: Asset Writedown
    r = 22
    write_label(ws, r, 1, "Asset Writedown", indent=2)
    write_crossref_or_zero(r, input_sheet, is_map, 'asset_writedown')

    # Row 23: Other Unusual Items
    r = 23
    write_label(ws, r, 1, "Other Unusual Items", indent=2)
    # Combine restructuring + other_unusual + merger_restruct if present
    _unusual_parts = []
    for _ukey in ['other_unusual', 'restructuring', 'merger_restruct', 'insurance_settlements']:
        if _ukey in is_map:
            _unusual_parts.append((_ukey, is_map[_ukey]))
    if len(_unusual_parts) == 0:
        write_zero_row(r)
    elif len(_unusual_parts) == 1:
        write_crossref_row(r, input_sheet, _unusual_parts[0][1])
    else:
        # Sum multiple unusual items
        for ci in range(2, 8):
            cl = get_column_letter(ci)
            cell = ws.cell(row=r, column=ci)
            parts = [f"IF(ISNUMBER({input_sheet}!{cl}{src_row}),{input_sheet}!{cl}{src_row},0)"
                     for _, src_row in _unusual_parts]
            cell.value = "=" + "+".join(parts)
            style_crossref_cell(cell, FMT_CURRENCY)

    # Row 24: Total Non-recurring = SUM(20:23)
    r = 24
    write_label(ws, r, 1, "Total Non-recurring", bold=True)
    write_formula_row(r, lambda cl: f"=SUM({cl}20:{cl}23)", bold=True)

    # Row 25: EBIT = Recurring Op Profit + Non-recurring
    r = 25
    write_label(ws, r, 1, "EBIT", bold=True)
    write_formula_row(r, lambda cl: f"={cl}17+{cl}24", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 26: blank
    r = 26

    # Row 27: Investment Income (Equity Method)
    r = 27
    write_label(ws, r, 1, "Investment Income (Equity Method)", indent=2)
    write_crossref_or_zero(r, input_sheet, is_map, 'income_from_affiliates')

    # Row 28: Interest Income
    r = 28
    write_label(ws, r, 1, "Interest Income", indent=2)
    write_crossref_row(r, input_sheet, is_map['interest_income'])

    # Row 29: Interest Expense
    r = 29
    write_label(ws, r, 1, "Interest Expense", indent=2)
    write_crossref_row(r, input_sheet, is_map['interest_expense'])

    # Row 30: FX Gains/(Losses)
    r = 30
    write_label(ws, r, 1, "FX Gains/(Losses)", indent=2)
    write_crossref_row(r, input_sheet, is_map['fx_gains'])

    # Row 31: Profit Before Tax = EBIT + 27 + 28 + 29 + 30
    r = 31
    write_label(ws, r, 1, "Profit Before Tax", bold=True)
    write_formula_row(r, lambda cl: f"={cl}25+{cl}27+{cl}28+{cl}29+{cl}30", bold=True)

    # Row 32: Tax Expense
    r = 32
    write_label(ws, r, 1, "Tax Expense", indent=2)
    write_crossref_row(r, input_sheet, is_map['tax_expense'])

    # Row 33: Net Income = PBT - Tax
    r = 33
    write_label(ws, r, 1, "Net Income", bold=True)
    write_formula_row(r, lambda cl: f"={cl}31-{cl}32", bold=True)
    style_double_line_row(ws, r, MAX_COL)

    # Row 34: blank
    r = 34

    # Row 35: Effective Tax Rate
    r = 35
    write_label(ws, r, 1, "Effective Tax Rate")
    write_formula_row(r, lambda cl: f"=IF({cl}31<>0,{cl}32/{cl}31,0)", fmt=FMT_PERCENT)

    # Row 36: NOPAT = Recurring Op Profit * (1 - ETR)
    r = 36
    write_label(ws, r, 1, "NOPAT", bold=True)
    write_formula_row(r, lambda cl: f"={cl}17*(1-{cl}35)", bold=True)

    # Row 37: blank
    r = 37

    # Row 38: Shares Outstanding (diluted)
    r = 38
    write_label(ws, r, 1, "Shares Outstanding (diluted)")
    write_crossref_row(r, input_sheet, is_map['diluted_shares'])

    # Row 39: EPS (diluted)  -- NI in thousands, shares in actuals
    r = 39
    write_label(ws, r, 1, "EPS (diluted)")
    write_formula_row(r, lambda cl: f"=IF({cl}38<>0,{cl}33/{cl}38*1000,0)", fmt='#,##0.00;(#,##0.00);"-"')

    # Row 40: Dividends per Share
    r = 40
    write_label(ws, r, 1, "Dividends per Share")
    write_crossref_or_zero(r, input_sheet, is_map, 'dividends_per_share', fmt='#,##0.00;(#,##0.00);"-"')

    # Row 41-42: blank
    r = 41
    r = 42

    # =========================================================================
    # Section 3: BALANCE SHEET (row 43 onwards)
    #
    # Uses CapIQ standardized data for consistency.
    # =========================================================================
    r = 43
    style_section_header(ws, r, MAX_COL, "BALANCE SHEET (Reformulated)")

    # --- Operating Working Capital ---
    r = 44
    write_label(ws, r, 1, "Operating Working Capital:", bold=True)

    # Row 45: Trade Receivables (CapIQ: includes contract assets)
    r = 45
    write_label(ws, r, 1, "Trade Receivables", indent=2)
    write_crossref_row(r, input_sheet, bs_map['accounts_receivable'])

    # Row 46: + Other Receivables (Income Tax Receivable etc.)
    r = 46
    write_label(ws, r, 1, "+ Other Receivables", indent=2)
    write_crossref_or_zero(r, input_sheet, bs_map, 'other_receivables')

    # Row 47: + Inventories
    r = 47
    write_label(ws, r, 1, "+ Inventories", indent=2)
    write_crossref_row(r, input_sheet, bs_map['inventory'])

    # Row 48: + Prepaid & Other CA
    r = 48
    write_label(ws, r, 1, "+ Prepaid & Other CA", indent=2)
    for ci in range(2, 8):
        cl = get_column_letter(ci)
        cell = ws.cell(row=r, column=ci)
        prep_row = bs_map['prepaid_exp']
        oca_row = bs_map['other_current_assets']
        cell.value = (
            f"=IF(ISNUMBER({input_sheet}!{cl}{prep_row}),{input_sheet}!{cl}{prep_row},0)"
            f"+IF(ISNUMBER({input_sheet}!{cl}{oca_row}),{input_sheet}!{cl}{oca_row},0)"
        )
        style_crossref_cell(cell, FMT_CURRENCY)

    # Row 49: - Accounts Payable
    r = 49
    write_label(ws, r, 1, "- Accounts Payable", indent=2)
    write_crossref_row(r, input_sheet, bs_map['accounts_payable'])

    # Row 50: - Accrued Expenses
    r = 50
    write_label(ws, r, 1, "- Accrued Expenses", indent=2)
    write_crossref_row(r, input_sheet, bs_map['accrued_exp'])

    # Row 51: - Unearned Revenue / Contract Liabilities
    r = 51
    write_label(ws, r, 1, "- Unearned Revenue / Contract Liabilities", indent=2)
    write_crossref_row(r, input_sheet, bs_map['unearned_revenue_curr'])

    # Row 52: - Income Taxes Payable
    r = 52
    write_label(ws, r, 1, "- Income Taxes Payable", indent=2)
    write_crossref_row(r, input_sheet, bs_map['current_income_tax'])

    # Row 53: - Other Current Liabilities
    r = 53
    write_label(ws, r, 1, "- Other Current Liabilities", indent=2)
    write_crossref_row(r, input_sheet, bs_map['other_current_liab'])

    # Row 54: = Operating Working Capital
    OWC_ROW = 54
    r = OWC_ROW
    write_label(ws, r, 1, "= Operating Working Capital", bold=True)
    write_formula_row(r, lambda cl: f"={cl}45+{cl}46+{cl}47+{cl}48-{cl}49-{cl}50-{cl}51-{cl}52-{cl}53", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 55: blank
    r = 55

    # --- Net Non-Current Operating Assets ---
    r = 56
    write_label(ws, r, 1, "Net Non-Current Operating Assets:", bold=True)

    # Row 57: PP&E (net) - CapIQ net_ppe (= Reported PPE + ROU)
    r = 57
    write_label(ws, r, 1, "PP&E (net, incl. ROU)", indent=2)
    write_crossref_row(r, input_sheet, bs_map['net_ppe'])

    # Row 58: + Goodwill
    r = 58
    write_label(ws, r, 1, "+ Goodwill", indent=2)
    write_crossref_row(r, input_sheet, bs_map['goodwill'])

    # Row 59: + Other Intangible Assets
    r = 59
    write_label(ws, r, 1, "+ Other Intangible Assets", indent=2)
    write_crossref_row(r, input_sheet, bs_map['other_intangibles'])

    # Row 60: + Deferred Tax Assets
    r = 60
    write_label(ws, r, 1, "+ Deferred Tax Assets", indent=2)
    write_crossref_or_zero(r, input_sheet, bs_map, 'deferred_tax_assets')

    # Row 61: + Deferred Charges & Other LT Intangibles
    r = 61
    write_label(ws, r, 1, "+ Deferred Charges & LT Intangibles", indent=2)
    write_crossref_or_zero(r, input_sheet, bs_map, 'deferred_charges_lt')

    # Row 62: + Other Non-Current Assets
    r = 62
    write_label(ws, r, 1, "+ Other Non-Current Assets", indent=2)
    write_crossref_or_zero(r, input_sheet, bs_map, 'other_lt_assets')

    # Row 63: - Deferred Tax Liabilities
    r = 63
    write_label(ws, r, 1, "- Deferred Tax Liabilities", indent=2)
    write_crossref_or_zero(r, input_sheet, bs_map, 'deferred_tax_liab')

    # Row 64: - Other Non-Current Liabilities (incl. pension if applicable)
    r = 64
    write_label(ws, r, 1, "- Other Non-Current Liabilities", indent=2)
    if 'pension_post_retire' in bs_map:
        # Sum other_non_current_liab + pension_post_retire
        for ci in range(2, 8):
            cl = get_column_letter(ci)
            cell = ws.cell(row=r, column=ci)
            ncl_row = bs_map['other_non_current_liab']
            pen_row = bs_map['pension_post_retire']
            cell.value = (
                f"=IF(ISNUMBER({input_sheet}!{cl}{ncl_row}),{input_sheet}!{cl}{ncl_row},0)"
                f"+IF(ISNUMBER({input_sheet}!{cl}{pen_row}),{input_sheet}!{cl}{pen_row},0)"
            )
            style_crossref_cell(cell, FMT_CURRENCY)
    else:
        write_crossref_row(r, input_sheet, bs_map['other_non_current_liab'])

    # Note: Lease liabilities (non-current) are classified as financing,
    # not deducted from operating NCA.

    # Row 65: = Net Non-Current Operating Assets
    NET_NCA_ROW = 65
    r = NET_NCA_ROW
    write_label(ws, r, 1, "= Net Non-Current Operating Assets", bold=True)
    write_formula_row(r, lambda cl: f"={cl}57+{cl}58+{cl}59+{cl}60+{cl}61+{cl}62-{cl}63-{cl}64", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 66: blank
    r = 66

    # Row 67: = Net Operating Assets = OWC + Net NCA
    NOA_ROW = 67
    r = NOA_ROW
    write_label(ws, r, 1, "= Net Operating Assets", bold=True)
    write_formula_row(r, lambda cl: f"={cl}{OWC_ROW}+{cl}{NET_NCA_ROW}", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 68: blank
    r = 68

    # --- Investment Assets ---
    r = 69
    write_label(ws, r, 1, "Investment Assets:", bold=True)

    # Row 70: Long-Term Investments
    r = 70
    write_label(ws, r, 1, "Long-Term Investments", indent=2)
    write_crossref_or_zero(r, input_sheet, bs_map, 'lt_investments')

    # Row 71: = Total Investment Assets (single line item here)
    INVEST_ASSETS_ROW = 71
    r = INVEST_ASSETS_ROW
    write_label(ws, r, 1, "= Total Investment Assets", bold=True)
    write_formula_row(r, lambda cl: f"={cl}70", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 72: blank
    r = 72

    # Row 73: = Business Assets = NOA + Investment Assets
    BIZ_ASSETS_ROW = 73
    r = BIZ_ASSETS_ROW
    write_label(ws, r, 1, "= Business Assets", bold=True)
    write_formula_row(r, lambda cl: f"={cl}{NOA_ROW}+{cl}{INVEST_ASSETS_ROW}", bold=True)
    style_double_line_row(ws, r, MAX_COL)

    # Row 74: blank
    r = 74

    # --- Financing ---
    r = 75
    style_section_header(ws, r, MAX_COL, "FINANCING")

    # Row 76: Debt items -- handle companies with explicit LT debt vs lease-only
    r = 76
    write_label(ws, r, 1, "Lease Liabilities (current)", indent=2)
    write_crossref_or_zero(r, input_sheet, bs_map, 'current_lease_liab')

    # Row 77: + Lease Liabilities (non-current) / LT Debt
    r = 77
    write_label(ws, r, 1, "+ Lease Liabilities (non-current)", indent=2)
    write_crossref_or_zero(r, input_sheet, bs_map, 'lt_leases')

    # Row 78: = Total Debt
    TOTAL_DEBT_ROW = 78
    r = TOTAL_DEBT_ROW
    write_label(ws, r, 1, "= Total Debt", bold=True)
    # For companies with explicit debt (lt_debt, current_lt_debt, st_borrowings),
    # use the CapIQ total_debt figure; for ASM/AIXA (lease-only), sum row 76+77.
    if 'total_debt' in bs_map and ('lt_debt' in bs_map or 'st_borrowings' in bs_map):
        # Company has financial debt beyond leases -- use CapIQ total_debt
        write_crossref_row(r, input_sheet, bs_map['total_debt'])
    else:
        # Lease-only debt (ASM, AIXA) -- sum current + non-current leases
        write_formula_row(r, lambda cl: f"={cl}76+{cl}77", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 79: Cash & Equivalents
    r = 79
    write_label(ws, r, 1, "Cash & Equivalents", indent=2)
    write_crossref_row(r, input_sheet, bs_map['cash_st_investments'])

    # Row 80: = Net Debt = Debt - Cash
    NET_DEBT_ROW = 80
    r = NET_DEBT_ROW
    write_label(ws, r, 1, "= Net Debt", bold=True)
    write_formula_row(r, lambda cl: f"={cl}{TOTAL_DEBT_ROW}-{cl}79", bold=True)

    # Row 81: Group Equity
    EQUITY_ROW = 81
    r = EQUITY_ROW
    write_label(ws, r, 1, "Group Equity")
    write_crossref_row(r, input_sheet, bs_map['total_equity'])

    # Row 82: = Invested Capital = Net Debt + Equity
    # (Business Assets = NOA + Investment = Net Debt + Equity)
    INVESTED_CAP_ROW = 82
    r = INVESTED_CAP_ROW
    write_label(ws, r, 1, "= Invested Capital", bold=True)
    write_formula_row(r, lambda cl: f"={cl}{NET_DEBT_ROW}+{cl}{EQUITY_ROW}", bold=True)
    style_double_line_row(ws, r, MAX_COL)

    # Row 83: BS Check (should = 0)
    BS_CHECK_ROW = 83
    r = BS_CHECK_ROW
    write_label(ws, r, 1, "BS Check (should = 0)", bold=True)
    write_formula_row(r, lambda cl: f"={cl}{BIZ_ASSETS_ROW}-{cl}{INVESTED_CAP_ROW}", bold=True)

    # Row 84-85: blank
    r = 84
    r = 85

    # =========================================================================
    # Section 4: CASH FLOW (row 86 onwards)
    # =========================================================================
    r = 86
    style_section_header(ws, r, MAX_COL, "FREE CASH FLOW")

    # Row 87: NOPAT (reference to IS NOPAT row 36)
    NOPAT_CF_ROW = 87
    r = NOPAT_CF_ROW
    write_label(ws, r, 1, "NOPAT")
    write_formula_row(r, lambda cl: f"={cl}36")

    # Row 88: + D&A
    DA_CF_ROW = 88
    r = DA_CF_ROW
    write_label(ws, r, 1, "+ D&A")
    write_crossref_row(r, input_sheet, cf_map['da_total'])

    # Row 89: - Change in OWC (OWC(t) - OWC(t-1)), 2019 = 0
    CHG_OWC_ROW = 89
    r = CHG_OWC_ROW
    write_label(ws, r, 1, "- Change in OWC")
    # Col B (2019) = 0
    cell = ws.cell(row=r, column=2, value=0)
    style_formula_cell(cell, FMT_CURRENCY)
    # Cols C-G = OWC(t) - OWC(t-1)
    for ci in range(3, 8):
        cl = get_column_letter(ci)
        prev_cl = get_column_letter(ci - 1)
        cell = ws.cell(row=r, column=ci)
        cell.value = f"={cl}{OWC_ROW}-{prev_cl}{OWC_ROW}"
        style_formula_cell(cell, FMT_CURRENCY)

    # Row 90: - CapEx (absolute value -- source is negative)
    CAPEX_CF_ROW = 90
    r = CAPEX_CF_ROW
    write_label(ws, r, 1, "- CapEx")
    for ci in range(2, 8):
        cl = get_column_letter(ci)
        cell = ws.cell(row=r, column=ci)
        src_row = cf_map['capex']
        cell.value = f"=ABS(IF(ISNUMBER({input_sheet}!{cl}{src_row}),{input_sheet}!{cl}{src_row},0))"
        style_crossref_cell(cell, FMT_CURRENCY)

    # Row 91: - Purchase of Intangibles (absolute value)
    PURCH_INTANG_ROW = 91
    r = PURCH_INTANG_ROW
    write_label(ws, r, 1, "- Purchase of Intangibles")
    if 'purchase_intangibles' in cf_map:
        for ci in range(2, 8):
            cl = get_column_letter(ci)
            cell = ws.cell(row=r, column=ci)
            src_row = cf_map['purchase_intangibles']
            cell.value = f"=ABS(IF(ISNUMBER({input_sheet}!{cl}{src_row}),{input_sheet}!{cl}{src_row},0))"
            style_crossref_cell(cell, FMT_CURRENCY)
    else:
        write_zero_row(r)

    # Row 92: = Operating CF after investment
    OP_CF_ROW = 92
    r = OP_CF_ROW
    write_label(ws, r, 1, "= Operating CF after investment", bold=True)
    write_formula_row(r, lambda cl: f"={cl}{NOPAT_CF_ROW}+{cl}{DA_CF_ROW}-{cl}{CHG_OWC_ROW}-{cl}{CAPEX_CF_ROW}-{cl}{PURCH_INTANG_ROW}", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 93: blank
    r = 93

    # Row 94: + Net Investment Profit after tax = Invest Income * (1 - ETR)
    NET_INV_PROFIT_ROW = 94
    r = NET_INV_PROFIT_ROW
    write_label(ws, r, 1, "+ Net Investment Profit after tax")
    write_formula_row(r, lambda cl: f"={cl}27*(1-{cl}35)")

    # Row 95: - Change in Investment Assets = InvestAssets(t) - InvestAssets(t-1)
    CHG_INVEST_ROW = 95
    r = CHG_INVEST_ROW
    write_label(ws, r, 1, "- Change in Investment Assets")
    # Col B (2019) = 0
    cell = ws.cell(row=r, column=2, value=0)
    style_formula_cell(cell, FMT_CURRENCY)
    for ci in range(3, 8):
        cl = get_column_letter(ci)
        prev_cl = get_column_letter(ci - 1)
        cell = ws.cell(row=r, column=ci)
        cell.value = f"={cl}{INVEST_ASSETS_ROW}-{prev_cl}{INVEST_ASSETS_ROW}"
        style_formula_cell(cell, FMT_CURRENCY)

    # Row 96: = FCF to Debt & Equity
    FCF_DE_ROW = 96
    r = FCF_DE_ROW
    write_label(ws, r, 1, "= FCF to Debt & Equity", bold=True)
    write_formula_row(r, lambda cl: f"={cl}{OP_CF_ROW}+{cl}{NET_INV_PROFIT_ROW}-{cl}{CHG_INVEST_ROW}", bold=True)
    style_total_row(ws, r, MAX_COL)

    # Row 97: - Interest Expense after tax = Interest Exp * (1 - ETR)
    INT_AFTER_TAX_ROW = 97
    r = INT_AFTER_TAX_ROW
    write_label(ws, r, 1, "- Interest Expense after tax")
    # Interest expense is row 29 in IS; typically negative (expense)
    write_formula_row(r, lambda cl: f"={cl}29*(1-{cl}35)")

    # Row 98: + Change in Debt = Debt(t) - Debt(t-1)
    CHG_DEBT_ROW = 98
    r = CHG_DEBT_ROW
    write_label(ws, r, 1, "+ Change in Debt")
    cell = ws.cell(row=r, column=2, value=0)
    style_formula_cell(cell, FMT_CURRENCY)
    for ci in range(3, 8):
        cl = get_column_letter(ci)
        prev_cl = get_column_letter(ci - 1)
        cell = ws.cell(row=r, column=ci)
        cell.value = f"={cl}{TOTAL_DEBT_ROW}-{prev_cl}{TOTAL_DEBT_ROW}"
        style_formula_cell(cell, FMT_CURRENCY)

    # Row 99: = FCF to Equity
    FCF_EQ_ROW = 99
    r = FCF_EQ_ROW
    write_label(ws, r, 1, "= FCF to Equity", bold=True)
    write_formula_row(r, lambda cl: f"={cl}{FCF_DE_ROW}-{cl}{INT_AFTER_TAX_ROW}+{cl}{CHG_DEBT_ROW}", bold=True)
    style_double_line_row(ws, r, MAX_COL)

    # --- Freeze panes at B5 ---
    freeze_panes(ws, row=5, col=2)

    return ws


def build_standardized_asm(wb):
    """Build the 'Std ASM' sheet -- delegates to generic function."""
    return build_standardized_tab(wb, 'ASM', "'Input ASM'", "'Input ASM (Reported)'", 'Std ASM')


# ---------------------------------------------------------------------------
# Adjustments documentation tab
# ---------------------------------------------------------------------------
def build_adjustments_tab(wb):
    """Build the 'Adjustments' tab documenting accounting analysis performed.

    This is a text-only documentation tab -- no Excel formulas needed.
    """
    ws = wb.create_sheet(title="Adjustments")

    # Column widths
    set_column_widths(ws, widths={1: 5, 2: 30, 3: 80})

    # Row 1: Title
    r = 1
    cell = ws.cell(row=r, column=1, value="Accounting Analysis & Adjustments - ASM International NV")
    cell.font = Font(size=14, bold=True, color="000000")

    # Row 3: Section header
    r = 3
    cell = ws.cell(row=r, column=1, value="Areas Examined")
    cell.font = Font(size=12, bold=True, color="000000")

    # ----- Areas data -----
    areas = [
        {
            "name": "1. R&D Capitalization Policy",
            "description": "ASM capitalizes development expenditure (significant amounts).",
            "finding": "Compare with peers: all semi equipment companies capitalize similarly.",
            "decision": "No Adjustment",
            "justification": "Consistent with industry practice and peers.",
        },
        {
            "name": "2. Equity Method Investment (ASMPT)",
            "description": "~25% stake in ASM Pacific Technology.",
            "finding": "Already separated as Investment Income in standardized IS.",
            "decision": "No Adjustment",
            "justification": "No further adjustment needed.",
        },
        {
            "name": "3. FX Gains/Losses",
            "description": "Highly volatile: ranged from -23M (2020) to +45M (2024).",
            "finding": "Already classified below EBIT in standardized IS (non-operating).",
            "decision": "No Adjustment",
            "justification": "Already handled in standardization.",
        },
        {
            "name": "4. Stock-Based Compensation",
            "description": "EUR 42M in 2024, growing over time.",
            "finding": "Already included in operating expenses in CapIQ data.",
            "decision": "No Adjustment",
            "justification": "No adjustment needed.",
        },
        {
            "name": "5. Goodwill (LPE Acquisition 2022)",
            "description": "EUR 321M from LPE acquisition.",
            "finding": "No impairment charges taken.",
            "decision": "No Adjustment",
            "justification": "Monitor for future impairment risk.",
        },
        {
            "name": "6. Asset Writedowns & Unusual Items",
            "description": "Small and sporadic.",
            "finding": "Already classified as non-recurring in standardized IS.",
            "decision": "No Adjustment",
            "justification": "No adjustment needed.",
        },
    ]

    r = 5  # start writing areas
    for area in areas:
        # Area name (bold)
        cell = ws.cell(row=r, column=1, value=area["name"])
        cell.font = BLACK_BOLD

        r += 1
        ws.cell(row=r, column=2, value="Description:").font = BLACK_BOLD
        ws.cell(row=r, column=3, value=area["description"]).font = BLACK_FONT

        r += 1
        ws.cell(row=r, column=2, value="Finding:").font = BLACK_BOLD
        ws.cell(row=r, column=3, value=area["finding"]).font = BLACK_FONT

        r += 1
        ws.cell(row=r, column=2, value="Decision:").font = BLACK_BOLD
        ws.cell(row=r, column=3, value=area["decision"]).font = BLACK_FONT

        r += 1
        ws.cell(row=r, column=2, value="Justification:").font = BLACK_BOLD
        ws.cell(row=r, column=3, value=area["justification"]).font = BLACK_FONT

        r += 2  # blank row between areas

    # Summary
    r += 1
    cell = ws.cell(row=r, column=1, value="Summary")
    cell.font = Font(size=12, bold=True, color="000000")
    r += 1
    ws.cell(
        row=r,
        column=1,
        value="No material adjustments required. The standardization process already handles "
              "all analytical reclassifications. Adjusted tabs pass through from Standardized tabs.",
    ).font = BLACK_FONT

    return ws


# ---------------------------------------------------------------------------
# Adjusted statement tabs (pass-through from Standardized)
# ---------------------------------------------------------------------------
def build_adjusted_tab(wb, company_key, std_sheet_name, output_sheet_name):
    """Build an Adjusted tab that mirrors a Standardized tab via cross-sheet refs.

    Since no material adjustments are being made, every data cell is a simple
    cross-sheet reference to the corresponding Standardized tab cell.

    Parameters
    ----------
    wb : openpyxl.Workbook
    company_key : str
        Key into COMPANY_INFO (e.g. 'ASM').
    std_sheet_name : str
        Name of the source Standardized sheet (e.g. 'Std ASM').
    output_sheet_name : str
        Name for the new Adjusted sheet (e.g. 'Adj ASM').
    """
    src_ws = wb[std_sheet_name]
    ws = wb.create_sheet(title=output_sheet_name)

    MAX_COL = 7  # columns A-G

    # Column widths
    widths = {1: 45}
    for c in range(2, 8):
        widths[c] = 14
    set_column_widths(ws, widths=widths)

    # Iterate through every row/col in the Standardized sheet and mirror
    max_row = src_ws.max_row or 1

    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, MAX_COL + 1):
            src_cell = src_ws.cell(row=row_idx, column=col_idx)
            tgt_cell = ws.cell(row=row_idx, column=col_idx)

            if src_cell.value is None:
                # Preserve formatting for section header fills, borders, etc.
                if src_cell.fill and src_cell.fill.fill_type == "solid":
                    tgt_cell.fill = _copy.copy(src_cell.fill)
                if src_cell.border:
                    tgt_cell.border = _copy.copy(src_cell.border)
                continue

            if col_idx == 1:
                # Column A: copy label text with same font/formatting
                tgt_cell.value = src_cell.value
                if src_cell.font:
                    tgt_cell.font = _copy.copy(src_cell.font)
                if src_cell.alignment:
                    tgt_cell.alignment = _copy.copy(src_cell.alignment)
            else:
                # Data columns (B-G): write cross-sheet reference
                cl = get_column_letter(col_idx)
                tgt_cell.value = f"='{std_sheet_name}'!{cl}{row_idx}"
                style_crossref_cell(tgt_cell, FMT_CURRENCY)

                # Preserve special number formats from source
                if src_cell.number_format and src_cell.number_format != 'General':
                    tgt_cell.number_format = src_cell.number_format

                # Preserve bold from source (for total rows)
                if src_cell.font and src_cell.font.bold:
                    tgt_cell.font = Font(color="008000", bold=True)

            # Preserve fills (section headers, etc.) and borders
            if src_cell.fill and src_cell.fill.fill_type == "solid":
                tgt_cell.fill = _copy.copy(src_cell.fill)
            if src_cell.border and src_cell.border != Border():
                tgt_cell.border = _copy.copy(src_cell.border)

    # Freeze panes at B5 (same as standardized)
    freeze_panes(ws, row=5, col=2)

    return ws


# ---------------------------------------------------------------------------
# Ratio Definitions tab (reference / documentation)
# ---------------------------------------------------------------------------
def build_ratio_definitions(wb):
    """Build the 'Ratio Definitions' tab -- a text-only reference of ratio formulas."""
    ws = wb.create_sheet(title="Ratio Definitions")

    # Column widths
    set_column_widths(ws, widths={1: 5, 2: 40, 3: 70})

    # Title
    cell = ws.cell(row=1, column=1, value="Ratio Definitions & Formulas")
    cell.font = Font(size=14, bold=True, color="000000")

    sections = [
        ("Profitability", [
            ("ROE", "Net Income / Average Total Equity"),
            ("DuPont Decomposition", "ROS x Asset Turnover x Financial Leverage"),
            ("RNOA", "NOPAT Margin x NOA Turnover"),
            ("Gross Profit Margin", "Gross Profit / Revenue"),
            ("EBIT Margin", "EBIT / Revenue"),
            ("NOPAT Margin", "NOPAT / Revenue"),
        ]),
        ("Asset Management", [
            ("Receivables Turnover", "Revenue / Avg Trade Receivables"),
            ("Days Receivable", "365 / Receivables Turnover"),
            ("Inventory Turnover", "COGS / Avg Inventory"),
            ("Days Inventory", "365 / Inventory Turnover"),
            ("Payables Turnover", "COGS / Avg Trade Payables"),
            ("Days Payable", "365 / Payables Turnover"),
            ("Cash Conversion Cycle", "Days Inventory + Days Receivable - Days Payable"),
        ]),
        ("Liquidity", [
            ("OWC / Revenue", "Operating Working Capital / Revenue"),
            ("Cash / Total Debt", "Cash & Equivalents / Total Debt"),
        ]),
        ("Solvency", [
            ("Debt-to-Equity", "Total Debt / Total Equity"),
            ("Interest Coverage", "EBIT / |Interest Expense|"),
        ]),
        ("Growth", [
            ("Revenue Growth", "(Revenue_t / Revenue_{t-1}) - 1"),
            ("Sustainable Growth", "(1 - Payout Ratio) x ROE"),
        ]),
    ]

    r = 3
    for section_name, ratios in sections:
        # Section header
        cell = ws.cell(row=r, column=1, value=section_name)
        cell.font = Font(size=12, bold=True, color="000000")
        cell.border = Border(bottom=Side(style="medium"))
        ws.cell(row=r, column=2).border = Border(bottom=Side(style="medium"))
        ws.cell(row=r, column=3).border = Border(bottom=Side(style="medium"))
        r += 1

        for name, formula in ratios:
            ws.cell(row=r, column=2, value=name).font = BLACK_BOLD
            ws.cell(row=r, column=3, value=formula).font = BLACK_FONT
            r += 1

        r += 1  # blank row between sections

    return ws


# ---------------------------------------------------------------------------
# Ratio Analysis tab (per company)
# ---------------------------------------------------------------------------
def build_ratio_tab(wb, adj_sheet_name, output_sheet_name):
    """Build a ratio analysis tab with Excel formulas referencing an Adjusted tab.

    Parameters
    ----------
    wb : openpyxl.Workbook
    adj_sheet_name : str
        Name of the Adjusted sheet (e.g. 'Adj ASM').
    output_sheet_name : str
        Name for the output ratio sheet (e.g. 'Ratios ASM').
    """
    ws = wb.create_sheet(title=output_sheet_name)
    MAX_COL = 7  # A-G

    adj = f"'{adj_sheet_name}'"

    # Column widths
    widths = {1: 40}
    for c in range(2, 8):
        widths[c] = 14
    set_column_widths(ws, widths=widths)

    # --- Title ---
    ws.cell(row=1, column=1,
            value=f"Ratio Analysis - {adj_sheet_name.replace('Adj ', '')}").font = Font(
        size=14, bold=True, color="000000")
    ws.cell(row=2, column=1, value="(derived from Adjusted statements)").font = Font(
        size=10, italic=True, color="666666")

    # Year headers in row 3
    write_year_headers(ws, 3, 2, YEARS)

    # --- Helpers ---
    def _label(row, text, bold=False, indent=0):
        write_label(ws, row, 1, text, bold=bold, indent=indent)

    def _formula(row, col_idx, formula_str, fmt=FMT_PERCENT):
        cell = ws.cell(row=row, column=col_idx)
        cell.value = formula_str
        style_formula_cell(cell, fmt)

    def _na(row, col_idx, fmt=FMT_PERCENT):
        """Write N/A for the first year where averages are not computable."""
        cell = ws.cell(row=row, column=col_idx, value="N/A")
        cell.font = Font(color="999999", italic=True)
        if fmt:
            cell.number_format = fmt

    def _cols():
        """Return list of (col_index, col_letter, prev_col_letter_or_None)."""
        result = []
        for ci in range(2, 8):
            cl = get_column_letter(ci)
            prev_cl = get_column_letter(ci - 1) if ci > 2 else None
            result.append((ci, cl, prev_cl))
        return result

    # =========================================================================
    # Section: PROFITABILITY
    # =========================================================================
    r = 5
    style_section_header(ws, r, MAX_COL, "PROFITABILITY")

    # --- ROE ---
    r = 6
    _label(r, "ROE", bold=True)
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci)
        else:
            _formula(r, ci,
                     f"=IF(AVERAGE({adj}!{prev_cl}81,{adj}!{cl}81)=0,0,"
                     f"{adj}!{cl}33/AVERAGE({adj}!{prev_cl}81,{adj}!{cl}81))")

    # --- DuPont: Net Profit Margin (ROS) ---
    r = 7
    _label(r, "  Net Profit Margin (ROS)", indent=2)
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}33/{adj}!{cl}9)")

    # --- DuPont: Asset Turnover ---
    r = 8
    _label(r, "  x Asset Turnover", indent=2)
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci,
                     f"=IF(AVERAGE({adj}!{prev_cl}73,{adj}!{cl}73)=0,0,"
                     f"{adj}!{cl}9/AVERAGE({adj}!{prev_cl}73,{adj}!{cl}73))",
                     fmt=FMT_RATIO)

    # --- DuPont: ROA = ROS x Asset Turnover ---
    r = 9
    _label(r, "  = ROA", indent=2)
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci)
        else:
            _formula(r, ci, f"={cl}7*{cl}8")

    # --- DuPont: Financial Leverage ---
    r = 10
    _label(r, "  x Financial Leverage", indent=2)
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci,
                     f"=IF(AVERAGE({adj}!{prev_cl}81,{adj}!{cl}81)=0,0,"
                     f"AVERAGE({adj}!{prev_cl}73,{adj}!{cl}73)/AVERAGE({adj}!{prev_cl}81,{adj}!{cl}81))",
                     fmt=FMT_RATIO)

    # --- DuPont: ROE check ---
    r = 11
    _label(r, "  = ROE check (ROA x Leverage)", indent=2)
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci)
        else:
            _formula(r, ci, f"={cl}9*{cl}10")

    # --- RNOA ---
    r = 13
    _label(r, "RNOA", bold=True)
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci)
        else:
            _formula(r, ci,
                     f"=IF(AVERAGE({adj}!{prev_cl}67,{adj}!{cl}67)=0,0,"
                     f"{adj}!{cl}36/AVERAGE({adj}!{prev_cl}67,{adj}!{cl}67))")

    # --- NOPAT Margin (for RNOA decomposition) ---
    r = 14
    _label(r, "  NOPAT Margin", indent=2)
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}36/{adj}!{cl}9)")

    # --- NOA Turnover ---
    r = 15
    _label(r, "  x NOA Turnover", indent=2)
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci,
                     f"=IF(AVERAGE({adj}!{prev_cl}67,{adj}!{cl}67)=0,0,"
                     f"{adj}!{cl}9/AVERAGE({adj}!{prev_cl}67,{adj}!{cl}67))",
                     fmt=FMT_RATIO)

    # --- Gross Profit Margin ---
    r = 17
    _label(r, "Gross Profit Margin")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}11/{adj}!{cl}9)")

    # --- EBIT Margin ---
    r = 18
    _label(r, "EBIT Margin")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}25/{adj}!{cl}9)")

    # --- NOPAT Margin ---
    r = 19
    _label(r, "NOPAT Margin")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}36/{adj}!{cl}9)")

    # =========================================================================
    # Section: COMMON-SIZE INCOME STATEMENT
    # =========================================================================
    r = 21
    style_section_header(ws, r, MAX_COL, "COMMON-SIZE INCOME STATEMENT")

    r = 22
    _label(r, "Revenue")
    for ci in range(2, 8):
        _formula(r, ci, "=1", fmt=FMT_PERCENT)

    r = 23
    _label(r, "Cost of Sales / Revenue")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}10/{adj}!{cl}9)")

    r = 24
    _label(r, "SG&A / Revenue")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}12/{adj}!{cl}9)")

    r = 25
    _label(r, "R&D / Revenue")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}13/{adj}!{cl}9)")

    # =========================================================================
    # Section: ASSET MANAGEMENT
    # =========================================================================
    r = 27
    style_section_header(ws, r, MAX_COL, "ASSET MANAGEMENT")

    # --- Receivables Turnover ---
    r = 28
    _label(r, "Receivables Turnover")
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci,
                     f"=IF(AVERAGE({adj}!{prev_cl}45,{adj}!{cl}45)=0,0,"
                     f"{adj}!{cl}9/AVERAGE({adj}!{prev_cl}45,{adj}!{cl}45))",
                     fmt=FMT_RATIO)

    # --- Days Receivable ---
    r = 29
    _label(r, "Days Receivable")
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci, f"=IF({cl}28=0,0,365/{cl}28)", fmt=FMT_RATIO)

    # --- Inventory Turnover ---
    r = 30
    _label(r, "Inventory Turnover")
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci,
                     f"=IF(AVERAGE({adj}!{prev_cl}47,{adj}!{cl}47)=0,0,"
                     f"{adj}!{cl}10/AVERAGE({adj}!{prev_cl}47,{adj}!{cl}47))",
                     fmt=FMT_RATIO)

    # --- Days Inventory ---
    r = 31
    _label(r, "Days Inventory")
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci, f"=IF({cl}30=0,0,365/{cl}30)", fmt=FMT_RATIO)

    # --- Payables Turnover ---
    r = 32
    _label(r, "Payables Turnover")
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci,
                     f"=IF(AVERAGE({adj}!{prev_cl}49,{adj}!{cl}49)=0,0,"
                     f"{adj}!{cl}10/AVERAGE({adj}!{prev_cl}49,{adj}!{cl}49))",
                     fmt=FMT_RATIO)

    # --- Days Payable ---
    r = 33
    _label(r, "Days Payable")
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci, f"=IF({cl}32=0,0,365/{cl}32)", fmt=FMT_RATIO)

    # --- Cash Conversion Cycle ---
    r = 34
    _label(r, "Cash Conversion Cycle", bold=True)
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci, fmt=FMT_RATIO)
        else:
            _formula(r, ci, f"={cl}31+{cl}29-{cl}33", fmt=FMT_RATIO)

    # --- OWC / Revenue ---
    r = 36
    _label(r, "OWC / Revenue")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}54/{adj}!{cl}9)")

    # --- PP&E / Revenue ---
    r = 37
    _label(r, "PP&E / Revenue")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}9=0,0,{adj}!{cl}57/{adj}!{cl}9)")

    # =========================================================================
    # Section: LIQUIDITY & SOLVENCY
    # =========================================================================
    r = 39
    style_section_header(ws, r, MAX_COL, "LIQUIDITY & SOLVENCY")

    # --- Cash / Total Debt ---
    r = 40
    _label(r, "Cash / Total Debt")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}78=0,0,{adj}!{cl}79/{adj}!{cl}78)",
                 fmt=FMT_RATIO)

    # --- Debt-to-Equity ---
    r = 41
    _label(r, "Debt-to-Equity")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF({adj}!{cl}81=0,0,{adj}!{cl}78/{adj}!{cl}81)",
                 fmt=FMT_RATIO)

    # --- Interest Coverage ---
    r = 42
    _label(r, "Interest Coverage")
    for ci, cl, _ in _cols():
        _formula(r, ci,
                 f"=IF(ABS({adj}!{cl}29)=0,0,{adj}!{cl}25/ABS({adj}!{cl}29))",
                 fmt=FMT_RATIO)

    # =========================================================================
    # Section: GROWTH
    # =========================================================================
    r = 44
    style_section_header(ws, r, MAX_COL, "GROWTH")

    # --- Revenue Growth ---
    r = 45
    _label(r, "Revenue Growth")
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci)
        else:
            _formula(r, ci,
                     f"=IF({adj}!{prev_cl}9=0,0,{adj}!{cl}9/{adj}!{prev_cl}9-1)")

    # --- Sustainable Growth ---
    r = 46
    _label(r, "Sustainable Growth")
    for ci, cl, prev_cl in _cols():
        if prev_cl is None:
            _na(r, ci)
        else:
            # Payout = DPS * shares / NI
            # Sust Growth = (1 - payout) * ROE
            # Guard: if NI=0, show 0
            _formula(r, ci,
                     f"=IF({adj}!{cl}33=0,0,"
                     f"(1-{adj}!{cl}40*{adj}!{cl}38/1000/{adj}!{cl}33)*{cl}6)")

    # Freeze panes at B5
    freeze_panes(ws, row=5, col=2)

    return ws


# ---------------------------------------------------------------------------
# Ratio Comparison tab (peer side-by-side)
# ---------------------------------------------------------------------------
def build_ratio_comparison(wb):
    """Build the 'Ratio Comparison' tab pulling key ratios from all 4 ratio tabs."""
    ws = wb.create_sheet(title="Ratio Comparison")
    MAX_COL = 7  # A-G

    # Column widths
    widths = {1: 40}
    for c in range(2, 8):
        widths[c] = 14
    set_column_widths(ws, widths=widths)

    # Title
    ws.cell(row=1, column=1,
            value="Ratio Comparison - Peer Analysis").font = Font(
        size=14, bold=True, color="000000")
    ws.cell(row=2, column=1,
            value="(key ratios side-by-side)").font = Font(
        size=10, italic=True, color="666666")

    # Year headers in row 3
    write_year_headers(ws, 3, 2, YEARS)

    # Ratio tabs and their display names
    ratio_tabs = [
        ("Ratios ASM", "ASM"),
        ("Ratios AIXTRON", "AIXTRON"),
        ("Ratios AMAT", "AMAT"),
        ("Ratios LRCX", "LRCX"),
    ]

    # Key ratios to compare: (label, source_row, format)
    key_ratios = [
        ("ROE", 6, FMT_PERCENT),
        ("RNOA", 13, FMT_PERCENT),
        ("NOPAT Margin", 19, FMT_PERCENT),
        ("Gross Profit Margin", 17, FMT_PERCENT),
        ("Revenue Growth", 45, FMT_PERCENT),
        ("OWC / Revenue", 36, FMT_PERCENT),
        ("Days Receivable", 29, FMT_RATIO),
        ("Debt-to-Equity", 41, FMT_RATIO),
        ("Interest Coverage", 42, FMT_RATIO),
    ]

    r = 5
    for ratio_label, src_row, fmt in key_ratios:
        # Section label for this ratio
        style_section_header(ws, r, MAX_COL, ratio_label)
        r += 1

        for tab_name, company_label in ratio_tabs:
            write_label(ws, r, 1, f"  {company_label}", indent=2)
            for ci in range(2, 8):
                cl = get_column_letter(ci)
                cell = ws.cell(row=r, column=ci)
                cell.value = f"='{tab_name}'!{cl}{src_row}"
                style_crossref_cell(cell, fmt)
            r += 1

        r += 1  # blank row between ratio groups

    # Freeze panes at B5
    freeze_panes(ws, row=5, col=2)

    return ws


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    print("Loading source data...")
    source_wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)

    print("Creating output workbook...")
    wb = openpyxl.Workbook()

    print("Building Cover...")
    build_cover(wb)

    print("Building Input tabs...")
    build_input_tabs(wb, source_wb)

    print("Building Std ASM...")
    build_standardized_asm(wb)

    print("Building Std AIXTRON...")
    build_standardized_tab(wb, 'AIXA', "'Input AIXTRON'", None, 'Std AIXTRON')

    print("Building Std AMAT...")
    build_standardized_tab(wb, 'AMAT', "'Input AMAT'", None, 'Std AMAT')

    print("Building Std LRCX...")
    build_standardized_tab(wb, 'LRCX', "'Input LRCX'", None, 'Std LRCX')

    print("Building Adjustments tab...")
    build_adjustments_tab(wb)

    print("Building Adj ASM...")
    build_adjusted_tab(wb, 'ASM', 'Std ASM', 'Adj ASM')

    print("Building Adj AIXTRON...")
    build_adjusted_tab(wb, 'AIXA', 'Std AIXTRON', 'Adj AIXTRON')

    print("Building Adj AMAT...")
    build_adjusted_tab(wb, 'AMAT', 'Std AMAT', 'Adj AMAT')

    print("Building Adj LRCX...")
    build_adjusted_tab(wb, 'LRCX', 'Std LRCX', 'Adj LRCX')

    print("Building Ratio Definitions...")
    build_ratio_definitions(wb)

    print("Building Ratios ASM...")
    build_ratio_tab(wb, 'Adj ASM', 'Ratios ASM')

    print("Building Ratios AIXTRON...")
    build_ratio_tab(wb, 'Adj AIXTRON', 'Ratios AIXTRON')

    print("Building Ratios AMAT...")
    build_ratio_tab(wb, 'Adj AMAT', 'Ratios AMAT')

    print("Building Ratios LRCX...")
    build_ratio_tab(wb, 'Adj LRCX', 'Ratios LRCX')

    print("Building Ratio Comparison...")
    build_ratio_comparison(wb)

    # Ensure output directory exists
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

    print(f"Saving to {OUTPUT_FILE}...")
    wb.save(OUTPUT_FILE)
    print("Done!")

    # Print summary
    print(f"\nWorkbook contains {len(wb.sheetnames)} sheets:")
    for i, name in enumerate(wb.sheetnames, 1):
        print(f"  {i}. {name}")

    source_wb.close()
    wb.close()


if __name__ == "__main__":
    main()
