# ASM International Valuation Model - Design Spec

## Overview

Build a professional Excel valuation model for ASM International NV (ENXTAM:ASM) as part of LSE course AC444 Valuation and Security Analysis. The model is the analytical backbone behind a 10-minute group presentation worth 40% of the course grade.

**Deliverable:** Single `.xlsx` workbook built via Python/openpyxl with multiple tabs covering standardized financial statements, accounting adjustments, ratio analysis (company + 3 peers), forecasting, WACC, DCF, PVAOI, comparable company analysis, and sensitivity tables.

**Company:** ASM International NV - Dutch semiconductor equipment manufacturer specializing in atomic layer deposition (ALD) and epitaxy. Traded on Euronext Amsterdam. Reports in EUR thousands. FY ends December 31.

**Peers (3):**
- AIXTRON SE (XTRA:AIXA) - German, semiconductor deposition equipment, similar European niche focus
- Applied Materials Inc (NASDAQGS:AMAT) - US, industry leader, direct competitor in deposition
- Lam Research Corp (NASDAQGS:LRCX) - US, major semi equipment player

**Rationale for peer selection:** All are semiconductor equipment companies. AIXTRON is the closest European peer with similar technology focus. Applied Materials and Lam Research are industry leaders providing benchmark comparison. Tokyo Electron excluded due to March fiscal year-end complicating direct comparison.

**Data period:** 2019-2024 (6 fiscal years). Forecasts: 2025E-2030E.

## Architecture

### Project Structure

```
scripts/
  build_model.py      # Main builder - constructs all tabs sequentially
  formatting.py       # IB color coding, number formats, style helpers
  recalc.py          # Formula validation script
  data_maps.py       # Row/column mappings from source Excel to our model
output/
  ASM_Valuation_Model.xlsx
ERRORS.md             # Error log per prompt rules
```

### Build Approach

Single monolithic `build_model.py` that:
1. Opens the source data file (`ASM International and Comps data.xlsx`)
2. Creates a new workbook
3. Builds tabs sequentially (input -> standardized -> adjusted -> ratios -> forecast -> valuation)
4. Every calculated cell uses an Excel formula (never Python-computed values)
5. Cross-sheet references use `='Sheet Name'!CellRef` syntax
6. Applies IB formatting throughout

### Data Flow

```
Source Excel (CapIQ data)
    |
    v
Input Tabs (raw data preserved)
    |
    v
Standardized Tabs (recast into standard format)
    |
    v
Adjustments Tab (document + apply adjustments)
    |
    v
Adjusted Tabs (post-adjustment statements)
    |
    v
Ratio Tabs (full ratio analysis per company)
    |
    v
Ratio Comparison Tab (side-by-side peer comparison)
    |
    v
Forecast Tab (projected IS/BS/CF 2025E-2030E)
    |
    v
Valuation Tab (WACC, DCF, PVAOI, Comps, Sensitivity, Football Field)
```

## Tab Specifications

### Tab 1: Cover
- Title: "ASM International NV - Equity Report"
- Subtitle: "AC444 Valuation and Security Analysis"
- Company ticker, currency, date
- Table of contents linking to each section

### Tabs 2-5: Input Tabs
- **Input ASM** - Copy raw CapIQ data (IS rows 13-99, CF rows 102-153, BS rows 156-230)
- **Input AIXTRON** - Same from AIXTRON sheet
- **Input AMAT** - Same from Applied Materials sheet
- **Input LRCX** - Same from Lam Research sheet
- Also copy ASM (as reported) data to Input ASM tab
- Blue text for all values (hardcoded inputs)
- Preserve source labels and structure

### Tabs 6-9: Standardized Statements

Each company gets one tab with IS, BS, CF reformulated into a standard analytical format.

**Assumptions row:**
- Operating cash percentage: 2% of revenue (assumption for minimum cash needed for operations)
- Row tracking years: 2019=6, 2020=5, ... 2024=1

**Income Statement (standardized by function):**

| Row | Item | Source/Formula |
|-----|------|---------------|
| 1 | Revenue | From input |
| 2 | Cost of Sales (by function) | From input (sign: positive = expense) |
| 3 | SG&A (by function) | From input |
| 4 | R&D (by function) | From input |
| 5 | Depreciation & Amortization (by nature) | From input supplementals or as-reported |
| 6 | Other Operating Income/Expense, net | From input |
| 7 | **= Recurring Operating Expense** | =Sum(2:6) |
| 8 | **= Recurring Operating Profit** | =Revenue - Recurring OpEx |
| 9 | Non-recurring items | Asset writedowns, gains/losses on sales, other unusual |
| 10 | **= EBIT** | =Recurring OP + Non-recurring |
| 11 | Investment income (equity method) | From input (ASM: share in ASMPT) |
| 12 | Interest income | From input |
| 13 | Interest expense | From input |
| 14 | FX gains/losses | From input |
| 15 | Other non-operating | From input |
| 16 | **= Profit before tax** | =EBIT + items 11-15 |
| 17 | Tax expense | From input |
| 18 | **= Net Income** | =PBT - Tax |
| 19 | Effective tax rate | =Tax / PBT |
| 20 | **NOPAT** | =Recurring OP x (1 - tax rate) |

**Balance Sheet (reformulated):**

| Section | Items |
|---------|-------|
| **Operating Working Capital** | |
| + Trade receivables | From BS |
| + Contract assets | From BS (as-reported) |
| + Inventories | From BS |
| + Prepaid & other current assets | From BS |
| - Trade payables | From BS |
| - Accrued expenses | From BS |
| - Contract liabilities | From BS |
| - Income taxes payable | From BS |
| - Other current liabilities | From BS |
| - Provisions (warranty) | From BS (as-reported) |
| **= Operating Working Capital** | Sum |
| **Net Non-Current Operating Assets** | |
| + PP&E (net) | From BS |
| + Right-of-use assets | From BS (as-reported) |
| + Goodwill | From BS |
| + Other intangible assets | From BS |
| + Deferred charges (capitalized dev costs) | From BS |
| + Evaluation tools at customers | From BS (as-reported) |
| + Other non-current assets | From BS |
| + Employee benefit assets | From BS (as-reported) |
| - Deferred tax liabilities | From BS |
| - Lease liabilities (non-current) | From BS |
| - Contingent consideration | From BS (as-reported) |
| **= Net Non-Current Operating Assets** | Sum |
| **= Net Operating Assets** | =OWC + Net NCA |
| **Investment Assets** | |
| + Investment in associates (ASMPT) | From BS |
| + Other investments | From BS (as-reported) |
| **= Total Investment Assets** | Sum |
| **= Business Assets** | =NOA + Investment Assets |
| **Financing** | |
| Debt (lease liabilities current + non-current) | From BS |
| + Cash and equivalents | From BS |
| = Net Debt | =Debt - Cash |
| Group Equity | From BS |
| **= Invested Capital** | =Debt + Equity |
| **Check: Business Assets = Invested Capital** | Should be 0 |

**Cash Flow Statement (reformulated):**

| Row | Item |
|-----|------|
| NOPAT | From IS |
| - Change in OWC | From BS year-over-year |
| - CapEx (PP&E + intangibles + capitalized dev) | From CF |
| + Depreciation & Amortization add-back | From CF |
| = Operating CF after investment | |
| + Net investment profit after tax | Equity method income x (1-t) |
| - Change in investment assets | From BS |
| = Free Cash Flow to D&E | |
| - Interest expense after tax | Interest x (1-t) |
| + Change in debt | From BS |
| = Free Cash Flow to Equity | |

### Tab 10: Adjustments ASM

Document the accounting analysis and any adjustments made.

**Key areas to examine:**

1. **R&D Capitalization Policy**
   - ASM capitalizes development expenditure (EUR 166M in 2024)
   - Compare capitalization rate (capex/total R&D) vs peers
   - Compare amortization rates of intangibles vs peers
   - Decision: Document comparison, adjust only if materially different

2. **Equity Method Investment (ASMPT stake)**
   - ~25% stake, carrying value EUR 904M
   - Share of income: EUR 10M in 2024 (down from EUR 78M in 2022)
   - Dividend received: EUR 14M in 2024
   - Treatment: Separate as investment income in standardized statements

3. **FX Gains/Losses**
   - Highly volatile: -EUR 23M (2020) to +EUR 45M (2024)
   - Treatment: Classify as non-recurring in standardized IS

4. **Stock-Based Compensation**
   - EUR 42M in 2024, growing
   - Already in operating expenses (allocated to COGS/SG&A/R&D)
   - No adjustment needed, just note

5. **Goodwill (LPE Acquisition 2022)**
   - EUR 321M, no impairment taken
   - Monitor but no adjustment needed

6. **Evaluation Tools at Customers**
   - EUR 110M in 2024, classified as non-current
   - Unique to ASM's business model (demo equipment)
   - Classify as operating asset

For each area: show the analysis, state whether adjustment is made, justify the decision.

### Tabs 11-14: Adjusted Statements

Same structure as Standardized tabs but with adjustments applied. If no material adjustments are needed (likely outcome for ASM given clean accounting), these tabs mirror the standardized tabs with formulas referencing the standardized tabs plus any adjustment overlays.

### Tab 15: Ratio Definitions

List all ratios with their formulas (text-only reference sheet), matching the Sanofi example structure:

**Profitability:**
- ROE = (Net Income + Non-recurring after tax) / Avg Equity
- DuPont: ROS x Asset Turnover x Financial Leverage
- RNOA = NOPAT Margin x NOA Turnover
- Gross profit margin, EBITDA margin
- Common-size IS (all items as % of revenue)

**Asset Management:**
- OWC/Revenue, NCA/Revenue, PP&E/Revenue
- Receivables turnover, Days receivable
- Inventory turnover, Days inventory
- Payables turnover, Days payable
- Cash conversion cycle

**Liquidity:**
- Current ratio, Quick ratio, Cash ratio, Operating CF ratio

**Solvency:**
- Liabilities-to-equity, Debt-to-equity, Debt-to-invested capital
- Interest coverage (earnings and CF based)

**Growth:**
- Sustainable growth rate = (1 - payout ratio) x ROE

### Tabs 16-19: Ratio Tabs (one per company)

All ratios computed using Excel formulas referencing the Adjusted tabs. Years 2019-2024 (some ratios need averages so effective from 2020).

### Tab 20: Ratio Comparison

Side-by-side comparison of key ratios across all 4 companies. Time series for each key ratio.

### Tab 21: Forecast ASM

**Top section: Condensed Balance Sheet (actual + forecast)**

| Item | 2022 | 2023 | 2024 | 2025E | 2026E | 2027E | 2028E | 2029E | 2030E |
|------|------|------|------|-------|-------|-------|-------|-------|-------|
| Net Working Capital | ref | ref | ref | forecast | ... | ... | ... | ... | ... |
| + Net non-current operating assets | | | | | | | | | |
| = Net operating assets | | | | | | | | | |
| + Investment Assets | | | | | | | | | |
| = Business Assets | | | | | | | | | |
| Debt | | | | | | | | | |
| + Group Equity | | | | | | | | | |
| = Invested Capital | | | | | | | | | |

**Middle section: Condensed Income Statement**
- Revenue, Gross Profit, EBITDA, EBIT, NOPAT
- Investment profit, Interest expense, Net Income
- Free Cash Flow build-up

**Bottom section: Forecast Assumptions (BLUE inputs with YELLOW background)**
- Revenue growth rate
- Gross profit margin
- EBITDA margin / EBIT margin
- NOPAT margin
- NWC/Revenue
- NCA/Revenue
- Investment assets/Revenue
- Tax rate
- After-tax cost of debt
- Debt to capital

**Revenue forecast approach:**
- 2024 actual: EUR 2,933M
- Backlog: EUR 1,566M (provides near-term visibility)
- Key drivers: AI/advanced packaging demand, ALD adoption, China exposure
- Assume moderate growth tapering toward industry long-term averages
- Growth rates as blue inputs that can be changed

### Tab 22: Valuation

**Section 1: Cost of Capital (WACC)**

| Input | Value | Source |
|-------|-------|--------|
| Risk-free rate | ~2.5% | 10-year German Bund (blue input) |
| Equity beta | ~1.2-1.4 | Regression or CapIQ (blue input) |
| Equity risk premium | ~5.0% | Damodaran implied ERP (blue input) |
| Cost of equity | CAPM formula | =Rf + Beta x ERP |
| Cost of debt | From financials | =Interest expense / Avg debt |
| Tax rate | From financials | =Effective tax rate |
| D/V | Market-based | =Debt / (Debt + Market Cap) |
| E/V | Market-based | =1 - D/V |
| **WACC** | Formula | =E/V x Re + D/V x Rd x (1-t) |

**Section 2: FCF to Equity Valuation**
- FCFE for each forecast year (from Forecast tab)
- Terminal value = FCFE_final x (1+g) / (Re - g)
- Discount factors based on cost of equity
- PV of FCFE -> Equity Value -> / shares = Price

**Section 3: FCF to Debt & Equity Valuation (DCF)**
- FCFF for each forecast year
- Terminal value = FCFF_final x (1+g) / (WACC - g)
- Discount factors based on WACC
- PV of FCFF -> Enterprise Value - Net Debt = Equity Value -> / shares = Price

**Section 4: PV of Abnormal NOPAT (PVAOI)**
- NOPAT each year
- Beginning BV of NOA
- Normal earnings = WACC x Beginning NOA
- Abnormal NOPAT = NOPAT - Normal earnings
- Terminal value of abnormal NOPAT
- PV -> + Beginning NOA = Enterprise Value - Net Debt = Equity Value -> / shares = Price

**Section 5: PV of Abnormal Earnings (PVAE)**
- Net Income each year
- Beginning BV of Equity
- Normal earnings = Re x Beginning Equity
- Abnormal Earnings = NI - Normal earnings
- PV -> + Current BV Equity = Equity Value -> / shares = Price

**Section 6: Sensitivity Tables (2-way)**
For each valuation method: rows = WACC (or Re), columns = terminal growth rate. Cell values = implied share price.

**Section 7: Comparable Company Analysis**
- Peer multiples: EV/EBITDA, EV/EBIT from actual market data
- Apply median/average to ASM's metrics
- Bridge: EV -> - Net Debt -> Equity Value -> / shares = Price

**Section 8: Football Field**
- Summary showing price range from each method
- Min/Max/Median across all methods
- Current market price for comparison

## Formatting Specification

### Color Coding (IB Standard)
- **Blue text** (RGB: 0,0,255): All hardcoded inputs and assumptions
- **Black text** (RGB: 0,0,0): All formulas and calculations
- **Green text** (RGB: 0,128,0): Cross-sheet references
- **Yellow background** (RGB: 255,255,0): Key assumptions requiring attention

### Number Formats
- Years: Text format (prevents "2,024")
- Currency: `#,##0` with units in headers ("EUR 000")
- Percentages: `0.0%`
- Multiples: `0.0"x"`
- Negatives in parentheses: `#,##0;(#,##0);"-"`
- Zeros displayed as "-": `#,##0;(#,##0);"-"`

### Layout
- Headers bold, borders below
- Section headers with background shading
- Freeze panes on row labels + year headers
- Column widths auto-fit or set to ~14 for data columns
- Print area set for each tab

## Validation

After building each tab, run `scripts/recalc.py` which:
1. Opens the workbook with openpyxl
2. Checks for any cell containing `#REF!`, `#NAME?`, `#VALUE!`, `#DIV/0!`, `#NULL!`
3. Validates that balance sheet checks = 0
4. Reports any issues
5. Zero errors required before proceeding to next tab

## Key Constraints

1. **Every calculated cell must contain an Excel formula** - never compute in Python and write a value
2. **Cross-sheet references must be explicit** - `='Standardized ASM'!B20` not a Python-computed number
3. **All assumptions in blue text** - so users can change inputs and the model recalculates
4. **Source data preserved in Input tabs** - model references these, never modifies source
5. **Error log maintained** in ERRORS.md throughout the build process
