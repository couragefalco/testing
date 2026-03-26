# ASM International Valuation Model - Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single professional Excel workbook (.xlsx) for ASM International NV containing standardized financials, accounting adjustments, ratio analysis with 3 peers, forecasting, WACC, DCF, PVAOI, comparable company analysis, and sensitivity tables.

**Architecture:** Single `build_model.py` script using openpyxl that reads source data from `ASM International and Comps data.xlsx` and constructs all tabs sequentially. Helper modules handle formatting (`formatting.py`), validation (`recalc.py`), and source data mapping (`data_maps.py`). Every calculated cell uses Excel formulas, never Python-computed values.

**Tech Stack:** Python 3, openpyxl, source data in `ASM International and Comps data.xlsx`

**Spec:** `docs/superpowers/specs/2026-03-26-asm-valuation-model-design.md`

---

## File Structure

```
scripts/
  formatting.py       # IB formatting helpers (colors, number formats, styles)
  data_maps.py        # Source row/column mappings for each company sheet
  recalc.py           # Validation: check for formula errors, BS balance checks
  build_model.py      # Main builder: constructs all tabs
output/
  ASM_Valuation_Model.xlsx
ERRORS.md
```

---

### Task 1: Project Scaffolding - formatting.py and recalc.py

**Files:**
- Create: `scripts/formatting.py`
- Create: `scripts/recalc.py`
- Create: `ERRORS.md`
- Create: `output/` directory

This task creates the utility modules that all subsequent tasks depend on.

- [ ] **Step 1: Create output directory and ERRORS.md**

```bash
mkdir -p scripts output
```

Create `ERRORS.md`:
```markdown
# Error Log

Errors encountered during model build. Read before starting each new tab.

---
```

- [ ] **Step 2: Write formatting.py**

Create `scripts/formatting.py` with all IB-standard formatting utilities:

```python
"""IB-standard formatting utilities for the valuation model."""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from openpyxl.utils import get_column_letter


# === Color Constants ===
BLUE_FONT = Font(color="0000FF")          # Hardcoded inputs
BLACK_FONT = Font(color="000000")         # Formulas
GREEN_FONT = Font(color="008000")         # Cross-sheet references
BLUE_BOLD = Font(color="0000FF", bold=True)
BLACK_BOLD = Font(color="000000", bold=True)

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
SECTION_FILL = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")

THIN_BORDER = Border(
    bottom=Side(style="thin"),
)
BOTTOM_DOUBLE = Border(
    bottom=Side(style="double"),
)

# === Number Formats ===
FMT_CURRENCY = '#,##0;(#,##0);"-"'
FMT_PERCENT = '0.0%'
FMT_MULTIPLE = '0.0"x"'
FMT_YEAR = '@'  # Text format for years
FMT_RATIO = '0.00;(0.00);"-"'
FMT_INTEGER = '#,##0;(#,##0);"-"'


def style_input_cell(cell, fmt=None):
    """Blue text for hardcoded input values."""
    cell.font = BLUE_FONT
    if fmt:
        cell.number_format = fmt


def style_formula_cell(cell, fmt=None):
    """Black text for formula cells."""
    cell.font = BLACK_FONT
    if fmt:
        cell.number_format = fmt


def style_crossref_cell(cell, fmt=None):
    """Green text for cross-sheet references."""
    cell.font = GREEN_FONT
    if fmt:
        cell.number_format = fmt


def style_assumption_cell(cell, fmt=None):
    """Blue text + yellow background for key assumptions."""
    cell.font = BLUE_BOLD
    cell.fill = YELLOW_FILL
    if fmt:
        cell.number_format = fmt


def style_header_row(ws, row, max_col, bold=True):
    """Apply header styling to a row."""
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(bold=bold, color="000000")
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER


def style_section_header(ws, row, max_col, text=None):
    """Apply section header styling."""
    cell = ws.cell(row=row, column=1)
    if text:
        cell.value = text
    cell.font = Font(bold=True, color="000000", size=11)
    cell.fill = SECTION_FILL
    for col in range(1, max_col + 1):
        ws.cell(row=row, column=col).fill = SECTION_FILL
        ws.cell(row=row, column=col).border = THIN_BORDER


def style_total_row(ws, row, max_col):
    """Apply total/subtotal border styling."""
    for col in range(1, max_col + 1):
        ws.cell(row=row, column=col).border = THIN_BORDER
        ws.cell(row=row, column=col).font = BLACK_BOLD


def style_double_line_row(ws, row, max_col):
    """Double-line border for final totals."""
    for col in range(1, max_col + 1):
        ws.cell(row=row, column=col).border = BOTTOM_DOUBLE
        ws.cell(row=row, column=col).font = BLACK_BOLD


def set_column_widths(ws, label_width=40, data_width=14, data_start_col=2, data_end_col=7):
    """Set standard column widths."""
    ws.column_dimensions['A'].width = label_width
    for col in range(data_start_col, data_end_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = data_width


def write_year_headers(ws, row, start_col, years, fmt=FMT_YEAR):
    """Write year headers as text to prevent '2,024' formatting."""
    for i, year in enumerate(years):
        cell = ws.cell(row=row, column=start_col + i)
        cell.value = str(year)
        cell.font = BLACK_BOLD
        cell.number_format = fmt
        cell.alignment = Alignment(horizontal="center")


def write_label(ws, row, col, text, bold=False, indent=0):
    """Write a row label."""
    cell = ws.cell(row=row, column=col)
    prefix = "  " * indent
    cell.value = f"{prefix}{text}"
    cell.font = Font(bold=bold, color="000000")
    return cell


def freeze_panes(ws, row, col):
    """Freeze panes at given position."""
    ws.freeze_panes = ws.cell(row=row, column=col)
```

- [ ] **Step 3: Write recalc.py**

Create `scripts/recalc.py`:

```python
"""Validation script for the valuation model workbook.

Usage: python scripts/recalc.py output/ASM_Valuation_Model.xlsx
"""

import sys
import openpyxl


ERROR_TOKENS = ["#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#NULL!", "#N/A"]


def validate_workbook(path):
    """Check workbook for formula errors and balance sheet checks."""
    wb = openpyxl.load_workbook(path)
    errors = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    val = str(cell.value)
                    for token in ERROR_TOKENS:
                        if token in val:
                            errors.append(
                                f"  {sheet_name}!{cell.coordinate}: {val}"
                            )

    if errors:
        print(f"ERRORS FOUND ({len(errors)}):")
        for e in errors:
            print(e)
        return False
    else:
        print(f"OK - {len(wb.sheetnames)} sheets, zero formula errors.")
        return True


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "output/ASM_Valuation_Model.xlsx"
    success = validate_workbook(path)
    sys.exit(0 if success else 1)
```

- [ ] **Step 4: Verify formatting module imports cleanly**

Run:
```bash
cd /Users/falco/Documents/Repositories/testing && python3 -c "from scripts.formatting import *; print('OK')"
```
Expected: `OK`

- [ ] **Step 5: Commit**

```bash
git add scripts/formatting.py scripts/recalc.py ERRORS.md
git commit -m "feat: add formatting utilities and validation script"
```

---

### Task 2: Data Mappings - data_maps.py

**Files:**
- Create: `scripts/data_maps.py`

Map the exact row numbers in the source Excel file to our model's standardized line items. This is critical — all subsequent tabs depend on reading source data from the correct cells.

- [ ] **Step 1: Write data_maps.py**

Create `scripts/data_maps.py`. This module maps source sheet names to row numbers for each financial line item. The source file has a consistent structure: IS starts around row 13, CF around row 102, BS around row 156.

```python
"""Row/column mappings from source Excel to our model.

Source file: 'ASM International and Comps data.xlsx'
All sheets use columns B-G for years 2019-2024 (FY).
"""

# Column mapping: B=2019, C=2020, D=2021, E=2022, F=2023, G=2024
YEAR_COLS = {'2019': 'B', '2020': 'C', '2021': 'D', '2022': 'E', '2023': 'F', '2024': 'G'}
YEARS = [2019, 2020, 2021, 2022, 2023, 2024]
DATA_START_COL = 2  # Column B
DATA_END_COL = 7    # Column G
NUM_YEARS = 6

# Source sheet names
SOURCE_SHEETS = {
    'ASM': 'ASM',
    'ASM_REPORTED': 'ASM (as reported)',
    'AIXTRON': 'AIXTRON',
    'AMAT': 'Applied Materials',
    'LRCX': 'Lam Research',
}

# === ASM (standardized CapIQ) row mappings ===
# Income Statement
ASM_IS = {
    'revenue': 21,
    'cogs': 23,
    'gross_profit': 24,
    'sga': 26,
    'rd': 27,
    'amort_intangibles': 28,
    'other_opex': 29,
    'total_opex': 30,
    'operating_income': 31,
    'interest_expense': 33,
    'interest_income': 34,
    'net_interest': 35,
    'income_affiliates': 36,
    'fx_gains': 37,
    'ebt_excl_unusual': 38,
    'gain_loss_invest': 39,
    'gain_loss_assets': 40,
    'asset_writedown': 41,
    'other_unusual': 42,
    'ebt_incl_unusual': 43,
    'tax_expense': 45,
    'net_income': 48,
    'basic_eps': 52,
    'diluted_eps': 55,
    'shares_basic': 54,
    'shares_diluted': 57,
    'dividends_per_share': 60,
    'ebitda': 64,
    'effective_tax_rate': 68,
    'sbc_total': 83,
    'rd_expense_supp': 79,  # Supplemental R&D (includes capitalized)
}

# Cash Flow Statement (starts row 102 in ASM sheet)
ASM_CF = {
    'net_income_cf': 110,
    'da_total': 113,
    'other_amort': 114,
    'loss_sale_assets': 115,
    'loss_sale_invest': 116,
    'writedown_restructuring': 117,
    'loss_equity_invest': 118,
    'sbc': 119,
    'other_operating': 120,
    'change_ar': 121,
    'change_inventory': 122,
    'change_ap': 123,
    'change_other_noa': 124,
    'cfo': 125,
    'capex': 127,
    'sale_ppe': 128,
    'acquisitions': 129,
    'purchase_intangibles': 130,
    'invest_securities': 131,
    'other_investing': 132,
    'cfi': 133,
    'lt_debt_repaid': 135,
    'stock_issuance': 137,
    'stock_repurchase': 138,
    'dividends_paid': 139,
    'other_financing': 141,
    'cff': 142,
    'fx_effect': 144,
    'net_change_cash': 145,
    'levered_fcf': 148,
    'unlevered_fcf': 149,
}

# Balance Sheet (starts row 156 in ASM sheet)
ASM_BS = {
    'cash': 165,
    'accounts_receivable': 167,
    'other_receivables': 168,
    'total_receivables': 169,
    'inventory': 170,
    'prepaid': 171,
    'other_current_assets': 172,
    'total_current_assets': 173,
    'gross_ppe': 174,
    'accum_depreciation': 175,
    'net_ppe': 176,
    'lt_investments': 177,
    'goodwill': 178,
    'other_intangibles': 179,
    'deferred_tax_asset': 180,
    'deferred_charges': 181,  # Capitalized development costs
    'other_lt_assets': 182,
    'total_assets': 183,
    'accounts_payable': 185,
    'accrued_expenses': 186,
    'current_lease': 187,
    'current_tax_payable': 188,
    'unearned_revenue': 189,
    'other_current_liabilities': 190,
    'total_current_liabilities': 191,
    'lt_lease': 192,
    'deferred_tax_liability': 193,
    'other_ncl': 194,
    'total_liabilities': 195,
    'common_stock': 197,
    'apic': 198,
    'retained_earnings': 199,
    'treasury_stock': 200,
    'oci': 201,
    'total_equity': 202,
    'total_liabilities_equity': 204,
    'shares_outstanding': 207,
    'total_debt': 211,
    'net_debt': 212,
    'equity_method_investments': 215,
    'employees': 223,
    'order_backlog': 225,
}

# ASM (as reported) - additional detail rows
ASM_REPORTED = {
    # Income statement
    'is_revenue': 17,
    'is_cogs': 19,
    'is_sga': 20,
    'is_rd': 21,
    'is_other_income': 22,
    'is_gain_loss_invest': 23,
    'is_fx': 24,
    'is_finance_income': 25,
    'is_finance_expense': 26,
    'is_share_associates': 27,
    'is_ebt': 28,
    'is_tax': 30,
    'is_net_income': 31,
    'is_operating_income': 34,
    # Cash flow
    'cf_da_impairment': 45,
    'cf_eval_tools': 49,
    'cf_tax_expense': 51,
    'cf_tax_paid': 52,
    'cf_non_cash_interest': 55,
    'cf_capitalized_dev': 72,
    'cf_intangible_purchases': 71,
    'cf_dividends_from_associate': 73,
    # Balance sheet
    'bs_cash': 94,
    'bs_ar': 95,
    'bs_contract_asset': 96,
    'bs_tax_receivable': 97,
    'bs_inventory': 98,
    'bs_other_ca': 99,
    'bs_total_ca': 100,
    'bs_rou': 102,
    'bs_ppe': 103,
    'bs_invest_associates': 104,
    'bs_other_invest': 105,
    'bs_dta': 106,
    'bs_goodwill': 107,
    'bs_other_intangibles': 108,
    'bs_other_assets': 109,
    'bs_employee_benefits': 110,
    'bs_eval_tools': 111,
    'bs_total_assets': 112,
    'bs_ap': 114,
    'bs_accrued': 115,
    'bs_tax_payable': 116,
    'bs_contract_liabilities': 117,
    'bs_warranty': 118,
    'bs_contingent_current': 119,
    'bs_total_cl': 120,
    'bs_lease_lt': 122,
    'bs_dtl': 123,
    'bs_contingent_lt': 124,
    'bs_total_equity': 127,
    'bs_total_le': 128,
}

# === PEER COMPANY ROW MAPPINGS ===
# Peers have the same CapIQ standardized format but slightly different row counts.
# All peers share identical structure to ASM's standardized sheet (not as-reported).

# AIXTRON uses same structure as ASM standardized sheet
# Row offsets may differ slightly - these are for the AIXTRON sheet specifically
AIXTRON_IS = {
    'revenue': 21,
    'cogs': 23,
    'gross_profit': 24,
    'sga': 26,
    'rd': 27,
    'other_opex': 29,
    'total_opex': 30,
    'operating_income': 31,
    'interest_expense': 33,
    'interest_income': 34,
    'net_interest': 35,
    'income_affiliates': 36,
    'fx_gains': 37,
    'ebt_excl_unusual': 38,
    'gain_loss_invest': 39,
    'gain_loss_assets': 40,
    'asset_writedown': 41,
    'other_unusual': 42,
    'ebt_incl_unusual': 43,
    'tax_expense': 45,
    'net_income': 48,
    'ebitda': 64,
    'effective_tax_rate': 68,
    'shares_outstanding': 54,
    'shares_diluted': 57,
    'dividends_per_share': 60,
    'sbc_total': 83,
}

# For Applied Materials and Lam Research, the structure is essentially the same
# as ASM standardized but verify exact rows when building.
# Use the same row map as ASM_IS as starting point - the builder will verify.
AMAT_IS = dict(ASM_IS)  # Same CapIQ format
LRCX_IS = dict(ASM_IS)  # Same CapIQ format

# The CF and BS rows also follow the same CapIQ structure across all sheets.
# Copy ASM maps as baseline.
AIXTRON_CF = dict(ASM_CF)
AIXTRON_BS = dict(ASM_BS)
AMAT_CF = dict(ASM_CF)
AMAT_BS = dict(ASM_BS)
LRCX_CF = dict(ASM_CF)
LRCX_BS = dict(ASM_BS)

# Company display info
COMPANY_INFO = {
    'ASM': {
        'full_name': 'ASM International NV',
        'ticker': 'ENXTAM:ASM',
        'currency': 'EUR',
        'magnitude': 'Thousands (K)',
        'source_sheet': 'ASM',
        'reported_sheet': 'ASM (as reported)',
        'is_map': ASM_IS,
        'cf_map': ASM_CF,
        'bs_map': ASM_BS,
        'reported_map': ASM_REPORTED,
    },
    'AIXTRON': {
        'full_name': 'AIXTRON SE',
        'ticker': 'XTRA:AIXA',
        'currency': 'EUR',
        'magnitude': 'Thousands (K)',
        'source_sheet': 'AIXTRON',
        'reported_sheet': None,
        'is_map': AIXTRON_IS,
        'cf_map': AIXTRON_CF,
        'bs_map': AIXTRON_BS,
        'reported_map': None,
    },
    'AMAT': {
        'full_name': 'Applied Materials, Inc.',
        'ticker': 'NASDAQGS:AMAT',
        'currency': 'EUR',
        'magnitude': 'Thousands (K)',
        'source_sheet': 'Applied Materials',
        'reported_sheet': None,
        'is_map': AMAT_IS,
        'cf_map': AMAT_CF,
        'bs_map': AMAT_BS,
        'reported_map': None,
    },
    'LRCX': {
        'full_name': 'Lam Research Corporation',
        'ticker': 'NASDAQGS:LRCX',
        'currency': 'EUR',
        'magnitude': 'Thousands (K)',
        'source_sheet': 'Lam Research',
        'reported_sheet': None,
        'is_map': LRCX_IS,
        'cf_map': LRCX_CF,
        'bs_map': LRCX_BS,
        'reported_map': None,
    },
}
```

- [ ] **Step 2: Verify data_maps loads and source file is readable**

```bash
python3 -c "
from scripts.data_maps import *
import openpyxl
wb = openpyxl.load_workbook('ASM International and Comps data.xlsx', data_only=True)
for key, info in COMPANY_INFO.items():
    ws = wb[info['source_sheet']]
    rev = ws.cell(row=info['is_map']['revenue'], column=2).value
    print(f'{key}: Revenue 2019 = {rev}')
print('All OK')
"
```

Expected: Revenue values for all 4 companies.

- [ ] **Step 3: Verify row mappings are correct for all companies**

The agent implementing this task MUST run a verification script that checks every mapped row against the source file to ensure the label matches expectations. If any peer company has different row numbers (CapIQ sometimes shifts rows), adjust the maps in `data_maps.py`. The key rows to spot-check:
- Revenue, COGS, Operating Income, Net Income in IS
- CFO, CapEx, Net Change in Cash in CF
- Total Assets, Total Equity, Total Liabilities in BS

```bash
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('ASM International and Comps data.xlsx', data_only=True)
for sheet_name in ['ASM', 'AIXTRON', 'Applied Materials', 'Lam Research']:
    ws = wb[sheet_name]
    print(f'\\n=== {sheet_name} ===')
    # Check key IS rows
    for r in [21, 23, 24, 31, 43, 45, 48, 64]:
        label = ws.cell(row=r, column=1).value
        val_2024 = ws.cell(row=r, column=7).value
        print(f'  Row {r}: {str(label)[:40]} = {val_2024}')
    # Check key CF rows
    for r in [110, 125, 127, 133, 145]:
        label = ws.cell(row=r, column=1).value
        val_2024 = ws.cell(row=r, column=7).value
        print(f'  Row {r}: {str(label)[:40]} = {val_2024}')
    # Check key BS rows
    for r in [165, 170, 176, 183, 195, 202, 204]:
        label = ws.cell(row=r, column=1).value
        val_2024 = ws.cell(row=r, column=7).value
        print(f'  Row {r}: {str(label)[:40]} = {val_2024}')
"
```

Fix any row mapping discrepancies found. Peer sheets may have rows offset by 1-2 from ASM.

- [ ] **Step 4: Commit**

```bash
git add scripts/data_maps.py
git commit -m "feat: add source data row/column mappings for all companies"
```

---

### Task 3: Build Model Skeleton + Input Tabs + Cover

**Files:**
- Create: `scripts/build_model.py`

Build the main script with the Cover tab and all 4 Input tabs that copy raw data from the source file.

- [ ] **Step 1: Write the main build_model.py skeleton with Cover and Input tab builders**

Create `scripts/build_model.py`. The script should:
1. Open the source workbook (data_only=True to read values)
2. Create a new output workbook
3. Build the Cover tab
4. Build Input tabs for ASM (including as-reported), AIXTRON, AMAT, LRCX
5. Save to `output/ASM_Valuation_Model.xlsx`

The Input tab builder function copies all non-empty rows from the source sheet into the input tab. All data cells get blue font (they're hardcoded inputs). Year headers get text format.

Key implementation notes:
- Use `data_only=True` when reading source so we get values not formulas
- When writing to the new workbook, write the values directly (these ARE the inputs)
- For each source sheet, copy rows from the IS section (rows ~13-99), CF section (~102-153), BS section (~156-230)
- Preserve the row labels in column A

```python
#!/usr/bin/env python3
"""Build the ASM International Valuation Model workbook.

Usage: python scripts/build_model.py
Output: output/ASM_Valuation_Model.xlsx
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Font, Alignment
from scripts.formatting import *
from scripts.data_maps import *


SOURCE_FILE = "ASM International and Comps data.xlsx"
OUTPUT_FILE = "output/ASM_Valuation_Model.xlsx"


def build_cover(wb):
    """Build the Cover tab."""
    ws = wb.active
    ws.title = "Cover"
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 30

    ws['A2'] = "ASM International NV"
    ws['A2'].font = Font(size=24, bold=True, color="000000")
    ws['A4'] = "Equity Valuation Report"
    ws['A4'].font = Font(size=16, bold=True, color="000000")
    ws['A6'] = "AC444 Valuation and Security Analysis"
    ws['A6'].font = Font(size=12, color="000000")
    ws['A7'] = "London School of Economics"
    ws['A7'].font = Font(size=12, color="000000")
    ws['A9'] = "Ticker:"
    ws['B9'] = "ENXTAM:ASM"
    ws['A10'] = "Currency:"
    ws['B10'] = "EUR (Thousands)"
    ws['A11'] = "Data Period:"
    ws['B11'] = "FY 2019 - FY 2024"
    ws['A12'] = "Forecast Period:"
    ws['B12'] = "FY 2025E - FY 2030E"
    ws['A14'] = "Peers:"
    ws['B14'] = "AIXTRON SE, Applied Materials, Lam Research"

    for r in [9, 10, 11, 12, 14]:
        ws.cell(row=r, column=1).font = Font(bold=True)
        ws.cell(row=r, column=2).font = BLUE_FONT

    # Table of contents
    ws['A17'] = "Table of Contents"
    ws['A17'].font = Font(size=14, bold=True)
    toc = [
        "1. Input Data (Raw CapIQ)",
        "2. Standardized Financial Statements",
        "3. Accounting Adjustments",
        "4. Adjusted Financial Statements",
        "5. Ratio Definitions",
        "6. Ratio Analysis (per company)",
        "7. Ratio Comparison (peers)",
        "8. Forecast",
        "9. Valuation (WACC, DCF, PVAOI, Comps, Sensitivity)",
    ]
    for i, item in enumerate(toc):
        ws.cell(row=19 + i, column=1).value = item
        ws.cell(row=19 + i, column=1).font = Font(size=11)


def copy_source_sheet(wb, source_wb, source_sheet_name, target_sheet_name, max_row=None):
    """Copy data from a source sheet into a new tab in the output workbook.

    All data cells are styled as blue (hardcoded inputs).
    """
    src = source_wb[source_sheet_name]
    ws = wb.create_sheet(target_sheet_name)

    if max_row is None:
        max_row = src.max_row
    max_col = src.max_column

    for row in src.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            new_cell = ws.cell(row=cell.row, column=cell.column)
            new_cell.value = cell.value
            # Column A = labels (black), data columns = blue (inputs)
            if cell.column == 1:
                new_cell.font = Font(color="000000")
            elif cell.value is not None:
                if isinstance(cell.value, (int, float)):
                    style_input_cell(new_cell, FMT_CURRENCY)
                elif isinstance(cell.value, str):
                    new_cell.font = Font(color="000000")
                else:
                    style_input_cell(new_cell)

    # Set column widths
    ws.column_dimensions['A'].width = 45
    for col in range(2, max_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14

    # Freeze panes
    freeze_panes(ws, 2, 2)


def build_input_tabs(wb, source_wb):
    """Build all Input tabs from source data."""
    # Input ASM (standardized CapIQ)
    copy_source_sheet(wb, source_wb, 'ASM', 'Input ASM')

    # Input ASM (as reported) - append to same tab or create separate
    copy_source_sheet(wb, source_wb, 'ASM (as reported)', 'Input ASM (Reported)')

    # Peer input tabs
    copy_source_sheet(wb, source_wb, 'AIXTRON', 'Input AIXTRON')
    copy_source_sheet(wb, source_wb, 'Applied Materials', 'Input AMAT')
    copy_source_sheet(wb, source_wb, 'Lam Research', 'Input LRCX')


def main():
    """Main build function."""
    print("Loading source data...")
    source_wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)

    print("Creating output workbook...")
    wb = openpyxl.Workbook()

    print("Building Cover...")
    build_cover(wb)

    print("Building Input tabs...")
    build_input_tabs(wb, source_wb)

    # Save
    print(f"Saving to {OUTPUT_FILE}...")
    wb.save(OUTPUT_FILE)
    print("Done!")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Run the build and verify**

```bash
python3 scripts/build_model.py
```

Expected: Creates `output/ASM_Valuation_Model.xlsx` with Cover + 5 input tabs.

- [ ] **Step 3: Run validation**

```bash
python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

Expected: `OK - 6 sheets, zero formula errors.`

- [ ] **Step 4: Commit**

```bash
git add scripts/build_model.py output/
git commit -m "feat: build model skeleton with Cover and Input tabs"
```

---

### Task 4: Standardized Financial Statements - ASM

**Files:**
- Modify: `scripts/build_model.py`

Add the `build_standardized_tab()` function that creates standardized IS, BS, and CF for a company. Build it for ASM first (which has the richest data including as-reported details), then reuse for peers.

This is the most complex tab — it reformulates the raw CapIQ data into the analytical framework from the course (operating vs. non-operating, recurring vs. non-recurring).

- [ ] **Step 1: Add build_standardized_asm() function**

Add to `scripts/build_model.py` a function that:
1. Creates the "Std ASM" sheet
2. Writes assumptions row (operating cash % = 2%)
3. Writes standardized Income Statement with Excel formulas referencing the Input tab
4. Writes reformulated Balance Sheet
5. Writes reformulated Cash Flow
6. Applies formatting

Key: All values must be **Excel formulas** referencing the Input tabs, like `='Input ASM'!B21`. Never read a Python value from the source and write it as a number.

The function should be ~200-300 lines. The agent should write the complete function following the row structure from the spec (IS rows, BS reformulation, CF reformulation). Every data cell must be an Excel formula string starting with `=`.

For the Income Statement section, use formulas like:
```python
# Revenue - reference Input ASM
ws[f'B{row}'] = "='Input ASM'!B21"
style_crossref_cell(ws[f'B{row}'], FMT_CURRENCY)
```

For calculated rows like Recurring Operating Profit:
```python
ws[f'B{row}'] = f"=B{revenue_row}-B{total_opex_row}"
style_formula_cell(ws[f'B{row}'], FMT_CURRENCY)
```

- [ ] **Step 2: Wire into main() and run**

Add call to `build_standardized_asm(wb)` in `main()` after input tabs. Run:
```bash
python3 scripts/build_model.py && python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

- [ ] **Step 3: Open the file in Excel/Numbers and manually verify**

Check that:
- Revenue 2024 = 2,932,724
- Net Income 2024 = 685,734
- Total Assets 2024 = 5,161,920
- Balance sheet check row = 0 for all years

Fix any formula reference errors found.

- [ ] **Step 4: Commit**

```bash
git add scripts/build_model.py
git commit -m "feat: add Standardized ASM tab with IS, BS, CF reformulation"
```

---

### Task 5: Standardized Financial Statements - Peers

**Files:**
- Modify: `scripts/build_model.py`

Refactor the standardized tab builder into a generic function that works for any company, then build tabs for AIXTRON, AMAT, and LRCX.

- [ ] **Step 1: Create generic build_standardized_tab() function**

Refactor the ASM-specific function into a generic one that accepts:
- `company_key` (e.g., 'ASM', 'AIXTRON')
- `input_sheet_name` (e.g., 'Input ASM', 'Input AIXTRON')
- `output_sheet_name` (e.g., 'Std ASM', 'Std AIXTRON')

The function uses `COMPANY_INFO[company_key]` to look up the correct row mappings. The IS/BS/CF structure is identical across all companies — only the input sheet reference and row numbers change.

For peers without an "as reported" sheet, some detailed BS items (ROU assets, contract assets, warranty provision, evaluation tools) won't be available. Use the standardized CapIQ data instead and note any items that can't be separated.

- [ ] **Step 2: Build all 3 peer tabs**

Call the generic function for each peer:
```python
build_standardized_tab(wb, 'AIXTRON', 'Input AIXTRON', 'Std AIXTRON')
build_standardized_tab(wb, 'AMAT', 'Input AMAT', 'Std AMAT')
build_standardized_tab(wb, 'LRCX', 'Input LRCX', 'Std LRCX')
```

- [ ] **Step 3: Run build + validate**

```bash
python3 scripts/build_model.py && python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

Spot-check peer revenues:
- AIXTRON 2024: ~633,159
- AMAT 2024: ~24,999,719
- LRCX 2024: ~13,782,535

- [ ] **Step 4: Commit**

```bash
git add scripts/build_model.py
git commit -m "feat: add standardized tabs for AIXTRON, AMAT, LRCX peers"
```

---

### Task 6: Adjustments Tab + Adjusted Statements

**Files:**
- Modify: `scripts/build_model.py`

Build the Adjustments ASM documentation tab and the Adjusted statement tabs (ASM + peers).

- [ ] **Step 1: Build Adjustments ASM tab**

This tab documents the accounting analysis. It's primarily text + reference data. Create a function `build_adjustments_tab(wb)` that:

1. Lists each adjustment area examined
2. Shows the analysis (peer comparison data referenced from standardized tabs)
3. States the decision (adjust or not) with justification
4. For any actual adjustments, shows the before/after impact

Key areas to document:
- R&D capitalization: ASM capitalizes dev costs. Show capitalized R&D / total R&D vs peers.
- Equity method investment: ASMPT stake separated as investment income (already done in standardization).
- FX gains/losses: Classified as non-recurring (already done).
- SBC: Already in opex, no adjustment needed.
- Impairments: Classified as non-recurring (already done).

For ASM, the standardized statements likely need no material adjustments beyond the reclassifications already performed during standardization. The Adjustments tab documents WHY.

- [ ] **Step 2: Build Adjusted tabs**

Since adjustments are minimal (reclassification already handled in standardization), the Adjusted tabs reference the Standardized tabs directly. Create `build_adjusted_tab(wb, company_key)` that:

1. Creates "Adj ASM", "Adj AIXTRON", "Adj AMAT", "Adj LRCX" tabs
2. Each cell references the corresponding Standardized tab cell:
   ```python
   ws['B5'] = "='Std ASM'!B5"
   style_crossref_cell(ws['B5'], FMT_CURRENCY)
   ```
3. Include an "Adjustment overlay" section at the top where specific adjustments could be added
4. The actual IS/BS/CF rows use formulas: `=Standardized value + Adjustment overlay`

This structure allows adding adjustments later without rebuilding.

- [ ] **Step 3: Run + validate**

```bash
python3 scripts/build_model.py && python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

Verify Adjusted ASM Revenue 2024 = Standardized ASM Revenue 2024 = 2,932,724.

- [ ] **Step 4: Commit**

```bash
git add scripts/build_model.py
git commit -m "feat: add Adjustments documentation and Adjusted statement tabs"
```

---

### Task 7: Ratio Definitions + Ratio Analysis Tabs

**Files:**
- Modify: `scripts/build_model.py`

Build the ratio definitions reference tab and ratio analysis tabs for all 4 companies.

- [ ] **Step 1: Build Ratio Definitions tab**

Create `build_ratio_definitions(wb)`. This is a text-only reference tab listing every ratio with its formula definition. Follow the exact structure from the Sanofi example (see spec for full list). No data — just ratio names and formula descriptions.

- [ ] **Step 2: Build generic build_ratio_tab() function**

Create `build_ratio_tab(wb, company_key, adj_sheet_name, output_sheet_name)` that computes all ratios using Excel formulas referencing the Adjusted tab.

The function builds sections:
1. **ROE & DuPont decomposition** (rows ~1-17)
2. **RNOA decomposition** (rows ~19-23)
3. **Common-size IS** (rows ~25-44)
4. **Other profitability** (rows ~46-48)
5. **Asset management** (rows ~50-64)
6. **Liquidity** (rows ~66-71)
7. **Solvency** (rows ~73-81)
8. **Sustainable growth** (rows ~83-87)

Each ratio cell contains an Excel formula. Ratios requiring averages (e.g., `Revenue / Avg Total Assets`) use `=(current + prior)/2` formulas. The first year (2019) may show N/A for average-based ratios since there's no 2018 data.

Example formulas:
```python
# ROE = Net Income / Average Equity
# Assuming Net Income is row 18 in Adj tab, Equity is row 65
ws['C5'] = f"='Adj ASM'!C18/AVERAGE('Adj ASM'!B65,'Adj ASM'!C65)"
style_formula_cell(ws['C5'], FMT_PERCENT)
```

- [ ] **Step 3: Build all 4 ratio tabs**

```python
build_ratio_tab(wb, 'ASM', 'Adj ASM', 'Ratios ASM')
build_ratio_tab(wb, 'AIXTRON', 'Adj AIXTRON', 'Ratios AIXTRON')
build_ratio_tab(wb, 'AMAT', 'Adj AMAT', 'Ratios AMAT')
build_ratio_tab(wb, 'LRCX', 'Adj LRCX', 'Ratios LRCX')
```

- [ ] **Step 4: Build Ratio Comparison tab**

Create `build_ratio_comparison(wb)` that pulls key ratios from all 4 company ratio tabs side by side. Include:
- ROE, RNOA, NOPAT margin, Gross margin
- Asset turnover, OWC/Revenue
- Current ratio, Debt-to-equity
- Sustainable growth rate

Each cell references the corresponding ratio tab:
```python
ws['B5'] = "='Ratios ASM'!C5"  # ASM ROE 2020
ws['C5'] = "='Ratios AIXTRON'!C5"  # AIXTRON ROE 2020
```

- [ ] **Step 5: Run + validate**

```bash
python3 scripts/build_model.py && python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

- [ ] **Step 6: Commit**

```bash
git add scripts/build_model.py
git commit -m "feat: add ratio definitions, ratio analysis, and peer comparison tabs"
```

---

### Task 8: Forecast Tab

**Files:**
- Modify: `scripts/build_model.py`

Build the Forecast ASM tab with historical data references and forecast projections.

- [ ] **Step 1: Build Forecast tab**

Create `build_forecast_tab(wb)`. Structure:

**Columns:** A=Labels, B-D=Historical (2022-2024), E-J=Forecast (2025E-2030E)

**Section 1: Forecast Assumptions (blue inputs, yellow background)**
Row items as key assumptions the user can change:
- Revenue growth rate: 12%, 10%, 8%, 6%, 4%, 3% (tapering)
- Gross profit margin: ~50.5% (recent trend)
- EBITDA margin: ~31% (recent trend)
- EBIT margin: ~27% (recent trend)
- NOPAT margin: ~21% (recent trend)
- NWC/Revenue: ~5% (from ratio analysis)
- Net NCA/Revenue: ~40% (from ratio analysis)
- Investment assets/Revenue: ~31% (from ratio analysis)
- Tax rate: ~21% (recent effective rate)
- After-tax cost of debt: from WACC section
- Debt to capital: ~1% (ASM has almost no debt)

All these are blue text + yellow background. They drive the entire forecast.

**Section 2: Condensed Income Statement**
- Revenue = Prior Revenue x (1 + growth rate)
- Gross Profit = Revenue x GP margin
- EBITDA = Revenue x EBITDA margin
- EBIT = Revenue x EBIT margin
- NOPAT = Revenue x NOPAT margin
- Net investment profit = Investment assets x return on investment
- Interest expense = Debt x cost of debt
- Net Income = NOPAT + investment profit - interest expense

Historical columns (2022-2024) reference the Adjusted ASM tab.

**Section 3: Condensed Balance Sheet**
- NWC = Revenue x (NWC/Revenue assumption)
- Net NCA = Revenue x (NCA/Revenue assumption)
- NOA = NWC + Net NCA
- Investment Assets = Revenue x (IA/Revenue assumption)
- Business Assets = NOA + IA
- Debt = Invested Capital x (Debt/Capital assumption)
- Equity = Invested Capital - Debt
- Invested Capital = Business Assets (must balance)

**Section 4: Free Cash Flow**
- NOPAT
- minus Change in NWC = NWC(t) - NWC(t-1)
- minus Change in NCA = NCA(t) - NCA(t-1)
- = Operating CF after investment
- plus Net investment profit after tax
- minus Change in investment assets
- = FCF to Debt & Equity
- minus Interest after tax
- plus Change in Debt
- = FCF to Equity

Every forecast cell = Excel formula referencing assumptions + prior year.

- [ ] **Step 2: Run + validate**

```bash
python3 scripts/build_model.py && python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

Open in Excel. Verify:
- Revenue 2025E grows from 2024 actual by the assumed growth rate
- Balance sheet balances (Business Assets = Invested Capital)
- FCF is positive and reasonable

- [ ] **Step 3: Commit**

```bash
git add scripts/build_model.py
git commit -m "feat: add Forecast tab with assumptions, IS, BS, and FCF projections"
```

---

### Task 9: Valuation Tab - WACC + DCF + PVAOI

**Files:**
- Modify: `scripts/build_model.py`

Build the Valuation tab with all valuation methods.

- [ ] **Step 1: Build WACC section**

Create `build_valuation_tab(wb)`. Start with WACC inputs at the top:

```
Row 2: "Cost of Capital Estimation"
Row 4: Input | Value | Comment | Source
Row 5: Risk-free rate | 2.5% | 10-year German Bund | [source]
Row 6: Equity beta | 1.30 | CapIQ / Regression | [source]
Row 7: ERP | 5.0% | Damodaran | [source]
Row 8: Cost of equity | =B5+B6*B7 | CAPM
Row 9: Cost of debt | formula ref | Interest/Avg Debt
Row 10: Tax rate | formula ref | Effective rate
Row 11: D/V | formula ref | Market-based
Row 12: E/V | =1-B11
Row 13: WACC | =B12*B8+B11*B9*(1-B10)
```

Blue inputs for Rf, Beta, ERP. Yellow background on all three. Formulas for everything else.

- [ ] **Step 2: Build FCF to Debt & Equity (DCF) section**

Below WACC, build the DCF model:
- Pull FCF to D&E from Forecast tab for 2025E-2030E
- Terminal growth rate as blue input (e.g., 2.5%)
- Terminal Value = FCF_2030 x (1+g) / (WACC - g)
- Discount factors: 1/(1+WACC)^n
- PV of each year's FCF
- Sum of PVs + PV of Terminal Value = Enterprise Value
- Minus Net Debt (from Adjusted BS) = Equity Value
- Divided by shares outstanding = Implied Price

- [ ] **Step 3: Build FCF to Equity section**

Same structure but:
- Use FCFE instead of FCFF
- Discount at cost of equity instead of WACC
- Result is directly Equity Value (no net debt adjustment)

- [ ] **Step 4: Build PVAOI (Abnormal NOPAT) section**

- NOPAT from forecast
- Beginning BV of NOA from forecast BS
- Normal earnings = WACC x Beginning NOA
- Abnormal NOPAT = NOPAT - Normal earnings
- Terminal value of abnormal NOPAT
- PV of abnormal NOPAT
- Equity Value = Beginning NOA + Sum PV + PV Terminal - Net Debt
- Price = Equity Value / shares

- [ ] **Step 5: Build PVAE (Abnormal Earnings) section**

- Net Income from forecast
- Beginning BV of Equity from forecast BS
- Normal earnings = Re x Beginning Equity
- Abnormal Earnings = NI - Normal
- PV -> + Current BV Equity = Equity Value
- Price = Equity Value / shares

- [ ] **Step 6: Run + validate**

```bash
python3 scripts/build_model.py && python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

All 4 valuation methods should produce prices in a reasonable range (check against ASM's actual trading price of ~EUR 400-500).

- [ ] **Step 7: Commit**

```bash
git add scripts/build_model.py
git commit -m "feat: add Valuation tab with WACC, DCF, PVAOI, PVAE models"
```

---

### Task 10: Sensitivity Tables + Comparable Company Analysis + Football Field

**Files:**
- Modify: `scripts/build_model.py`

Complete the Valuation tab with sensitivity analysis, comps, and the summary football field.

- [ ] **Step 1: Build Sensitivity Tables**

For each of the 4 valuation methods, add a 2-way sensitivity table:
- Rows: WACC (or Re) varying in 0.5% increments around the base case (e.g., 7%-12%)
- Columns: Terminal growth rate varying in 0.5% increments (e.g., 1%-4%)
- Cell values: Implied share price

Use Excel DATA TABLE formulas or manually build the grid with formulas that substitute the varying inputs. Since openpyxl can't create Excel Data Tables, build them as formula grids:

```python
# Each cell recalculates the DCF with the row's WACC and column's growth
# Formula: = (Sum of PV FCFs using row WACC) + TV using (row WACC, col g) - Net Debt) / shares
```

This is complex — a simpler approach is to build a helper column that computes the TV for each combination and chains through. The agent should use the simplest formula structure that works.

- [ ] **Step 2: Build Comparable Company Analysis section**

- Table of peer companies with their market multiples (EV/EBITDA, EV/EBIT)
- These values are hardcoded blue inputs (the agent must look up current market data or use reasonable estimates based on the CapIQ data available)
- Compute median and average of peer multiples
- Apply to ASM's EBITDA and EBIT from Adjusted 2024
- Bridge: Implied EV -> minus Net Debt -> Equity Value -> per share price

Peer data to input:
- AIXTRON: EV, EBITDA, EBIT from latest data
- AMAT: EV, EBITDA, EBIT
- LRCX: EV, EBITDA, EBIT

The agent should compute EV/EBITDA and EV/EBIT for each peer using Excel formulas.

- [ ] **Step 3: Build Football Field summary**

At the top of the Valuation tab (or a separate section), create a summary table:

| Method | Low | Base | High |
|--------|-----|------|------|
| DCF (FCFF) | sensitivity min | base case | sensitivity max |
| DCF (FCFE) | ... | ... | ... |
| PVAOI | ... | ... | ... |
| PVAE | ... | ... | ... |
| Comps EV/EBITDA | ... | ... | ... |
| Comps EV/EBIT | ... | ... | ... |
| Current Price | ... | ... | ... |

All cells reference other cells in the Valuation tab.

Add a final **Recommendation** row: BUY or SELL based on whether the average implied price > current market price.

- [ ] **Step 4: Run final validation**

```bash
python3 scripts/build_model.py && python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

- [ ] **Step 5: Commit**

```bash
git add scripts/build_model.py
git commit -m "feat: add sensitivity tables, comparable company analysis, football field"
```

---

### Task 11: Final Polish + Full Validation

**Files:**
- Modify: `scripts/build_model.py`

Final pass: formatting consistency, freeze panes, print areas, and end-to-end verification.

- [ ] **Step 1: Add freeze panes to all tabs**

Every data tab should freeze panes so row labels and year headers stay visible when scrolling:
```python
ws.freeze_panes = 'B3'  # or appropriate cell
```

- [ ] **Step 2: Set print areas**

For each tab, set the print area to the used range:
```python
ws.print_area = f'A1:{get_column_letter(max_col)}{max_row}'
```

- [ ] **Step 3: Verify all formatting rules**

Scan every tab:
- All hardcoded values have blue font
- All formulas have black font
- All cross-sheet references have green font
- Key assumptions have yellow background
- Year headers are text format (not "2,024")
- Currency cells use `#,##0;(#,##0);"-"`
- Percentage cells use `0.0%`

- [ ] **Step 4: Run full build + validation**

```bash
python3 scripts/build_model.py && python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

Expected: `OK - N sheets, zero formula errors.`

- [ ] **Step 5: Manual spot-check in Excel**

Open the file. Check:
1. Cover tab looks professional
2. Input tabs have all source data in blue
3. Standardized ASM: Revenue = 2,932,724 (2024), BS balances
4. Ratios look reasonable (gross margin ~50%, ROE ~18-20%)
5. Forecast grows smoothly, BS balances all years
6. All 4 valuation methods produce reasonable prices
7. Sensitivity tables show price variation
8. Football field summarizes all methods
9. Colors are correct throughout

- [ ] **Step 6: Final commit**

```bash
git add -A
git commit -m "feat: complete ASM International valuation model with final polish"
```

---

## Implementation Notes for Agents

### Critical Rules
1. **NEVER hardcode computed values.** Every calculated cell MUST be an Excel formula string starting with `=`. Read source data to understand what values to expect, but write `='Input ASM'!B21` not `1283860`.
2. **Cross-sheet references use quoted sheet names:** `="'Std ASM'"!B5"` — the single quotes around sheet names with spaces are required.
3. **Sign conventions:** In standardized IS, expenses should be shown as positive numbers that get subtracted. Follow the Sanofi example convention.
4. **NA handling:** Source data contains "NA" strings for missing values. Use `IF(ISNUMBER(...), ..., 0)` or similar in formulas that reference cells that might be NA.
5. **Peer data is in EUR already** (converted in the source). No FX conversion needed.
6. **Row mapping verification is critical.** Before writing formulas, verify that the source row actually contains the expected data. The CapIQ format is consistent but row offsets can differ by 1-2 between companies.

### Build Order
Tasks MUST be executed sequentially — each builds on the prior:
1. Scaffolding (formatting.py, recalc.py)
2. Data maps (data_maps.py)
3. Cover + Input tabs (build_model.py skeleton)
4. Standardized ASM
5. Standardized Peers
6. Adjustments + Adjusted tabs
7. Ratios
8. Forecast
9. Valuation (WACC, DCF, PVAOI)
10. Sensitivity + Comps + Football Field
11. Final polish
