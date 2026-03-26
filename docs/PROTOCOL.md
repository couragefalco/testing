# Full Protocol: Building the ASM International Valuation Model

## Project Overview

**Objective:** Build a professional Excel valuation workbook (.xlsx) for ASM International NV (ENXTAM:ASM) using Python/openpyxl, for the LSE course AC444 Valuation and Security Analysis.

**Final Deliverable:** `output/ASM_Valuation_Model.xlsx` — 23-tab workbook covering standardized financials, accounting adjustments, ratio analysis (company + 3 peers), forecasting, WACC, DCF, PVAOI, comparable company analysis, sensitivity tables, and a football field summary.

**Repo:** https://github.com/couragefalco/testing

---

## Phase 1: Project Setup & Data Exploration

### 1.1 Repository Creation
- Created GitHub repo `couragefalco/testing` using `gh repo create`
- Initialized with README, pushed to `main`

### 1.2 Unpacking Source Materials
- Unzipped `Claude Input.zip` containing course materials:
  - **Full Project Examples:** Sanofi and BioNTech Excel workbooks + presentation PDFs (used as structural reference)
  - **Criteria:** AC444 Marking Criteria, Course Outline
  - **Lecture Materials:** Slides on standardization, adjustments, ratios, DCF, PVAOI, cost of capital, continuing values, price multiples, strategy analysis, ESG
  - **Group Project Specification:** Detailed requirements document
- Read `Prompt Claude.docx` — the master prompt defining rules for building the model (Excel formulas only, IB color coding, error logging, validation after every tab)
- Read `ASM International and Comps data.xlsx` — source data file containing:
  - **ASM** sheet: Standardized CapIQ Income Statement, Cash Flow, Balance Sheet (2019-2024)
  - **ASM (as reported)** sheet: As-reported financials with additional detail (ROU assets, contract assets/liabilities, evaluation tools, warranty provisions, capitalized development expenditure)
  - **Electron** sheet: Tokyo Electron (excluded from final model due to March fiscal year-end)
  - **Applied Materials** sheet: Applied Materials Inc. (peer)
  - **Lam Research** sheet: Lam Research Corp. (peer)
  - **AIXTRON** sheet: AIXTRON SE (peer)
  - All peer data converted to EUR thousands in the source file

### 1.3 Studying Example Projects
- Read Sanofi Valuation Excel (AC444_S3G1) in detail:
  - 35+ tabs including: Cover, Valuation, Forecast, Input sheets, Standardized statements (per company), Adjusted statements (per company), Ratio Analysis (per company + definitions), Adjustments documentation, Comps, Beta regression
  - Key structural patterns: Condensed BS/IS in Forecast tab, Football Field summary, Sensitivity grids, Adjustment overlay structure
- Read BioNTech Excel (S1G1) — different structure but same core components:
  - Standardized -> Adjusted -> Ratios -> DCF -> PVAOI -> Comps pipeline
  - Revenue forecast by product segment
  - WACC with beta regression
- Read Sanofi and BioNTech presentation PDFs to understand quality expectations:
  - Sanofi: Highly polished, branded slides with strategy analysis, M&A timeline, Porter's Five Forces, ESG pillars, ratio charts, football field
  - BioNTech: Project timeline, challenges, condensed IS with annotations, product pipeline revenue model

### 1.4 Reading Course Requirements
- **Group Project Specification (5 pages):**
  - Required sections: (1) Strategy & competitiveness, (2) Financial statement & accounting analysis, (3) Ratio analysis, (4) Prospective analysis & forecasts, (5) Valuation (DCF + PVAOI) with buy/sell recommendation, (6) ESG factors
  - 10-minute presentation + 2-min Q&A
  - At least 3 years of data, 2-3 peer firms
  - Appendix: Original + recast statements, ratio calculations, all background work
- **Marking Criteria:** Over 70% (Distinction) requires "convincing command of course materials" and "original insights"
- **Course Outline:** Topics covered — method of comparables, accounting-based valuation, DCF, forecasting, cost of capital, continuing values, M&A valuation

---

## Phase 2: Design & Planning

### 2.1 Design Decisions

**Peer Selection (3 peers):**
- AIXTRON SE — European, same deposition equipment niche, closest comparable
- Applied Materials — US industry leader, direct competitor in deposition
- Lam Research — US major player, strong benchmark
- Tokyo Electron excluded: March fiscal year-end makes direct YoY comparison messy

**Architecture Decision:**
- Single `build_model.py` script (vs. script-per-tab) — chosen for cross-sheet reference integrity
- Helper modules for formatting (`formatting.py`), data mapping (`data_maps.py`), validation (`recalc.py`)
- All calculated cells use Excel formulas (never Python-computed values)

**Key Accounting Decisions:**
- R&D capitalization: No adjustment — ASM's policy is consistent with semiconductor equipment peers
- Equity method investment (ASMPT stake): Separated as investment income in standardized IS (already handled)
- FX gains/losses: Classified below EBIT as non-operating (already handled in standardization)
- Stock-based compensation: Already in operating expenses per CapIQ, no adjustment needed
- Goodwill from LPE acquisition: No impairment adjustment — EUR 321M stable since 2022
- Overall: No material adjustments beyond reclassifications performed during standardization

### 2.2 Design Spec
Saved to `docs/superpowers/specs/2026-03-26-asm-valuation-model-design.md`

Contents:
- Tab-by-tab specification (22 planned tabs, became 23)
- Standardized statement structure (IS, reformulated BS, CF)
- Ratio definitions (profitability, asset management, liquidity, solvency, growth)
- Forecast methodology (revenue growth tapering, margin assumptions)
- Valuation methods (DCF FCFF, DCF FCFE, PVAOI, PVAE, comps)
- Formatting specification (IB color coding, number formats)
- Validation approach

### 2.3 Implementation Plan
Saved to `docs/superpowers/plans/2026-03-26-asm-valuation-model.md`

11 sequential tasks:
1. Project scaffolding (formatting.py, recalc.py, ERRORS.md)
2. Data mappings (data_maps.py)
3. Model skeleton + Input tabs + Cover
4. Standardized ASM (IS, BS, CF reformulation)
5. Standardized Peers (generic function for AIXTRON, AMAT, LRCX)
6. Adjustments documentation + Adjusted tabs
7. Ratio definitions + Ratio analysis + Peer comparison
8. Forecast (assumptions, IS, BS, FCF projections)
9. Valuation (WACC, DCF, PVAOI, PVAE)
10. Sensitivity tables + Comps + Football Field
11. Final polish + full validation

---

## Phase 3: Implementation

### Execution Method: Subagent-Driven Development
Each task was dispatched to a fresh AI subagent with:
- Complete task description including exact row references, formula templates, and data map keys
- Context about prior tasks and the workbook structure so far
- Strict rules: Excel formulas only, IB color coding, commit after each task
- Post-task validation: `python3 scripts/build_model.py && python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx`

### Task 1: Project Scaffolding
**Subagent output:** DONE (95s)

Created:
- `scripts/__init__.py` — Package init
- `scripts/formatting.py` (198 lines) — Full IB formatting module:
  - 5 font constants: `BLUE_FONT` (0000FF), `BLACK_FONT` (000000), `GREEN_FONT` (008000), `BLUE_BOLD`, `BLACK_BOLD`
  - 3 fill constants: `YELLOW_FILL` (FFFF00), `HEADER_FILL` (D9E1F2), `SECTION_FILL` (B4C6E7)
  - 2 border constants: `THIN_BORDER`, `BOTTOM_DOUBLE`
  - 6 number format constants: `FMT_CURRENCY`, `FMT_PERCENT`, `FMT_MULTIPLE`, `FMT_YEAR`, `FMT_RATIO`, `FMT_INTEGER`
  - 12 helper functions for styling cells, rows, sections
- `scripts/recalc.py` (50 lines) — CLI validation script scanning for Excel error tokens
- `ERRORS.md` — Error log (empty, as no errors were encountered)
- `output/.gitkeep` — Output directory placeholder

**Verification:** `from scripts.formatting import *` — OK

### Task 2: Data Mappings
**Subagent output:** DONE (269s)

Created `scripts/data_maps.py` (896 lines) — THE most critical module.

The subagent:
1. Opened the source Excel file with openpyxl
2. Iterated through every row in every sheet to find the actual row numbers
3. Discovered significant row offset differences between companies:
   - ASM: IS starts row 21, CF row 110, BS row 165
   - AMAT: IS starts row 21, CF row 97(!), BS row 149(!)
   - LRCX: IS starts row 21, CF row 87(!), BS row 140(!)
   - AIXTRON: Generally similar to ASM but with some unique rows
4. Built separate row mapping dicts for each company: `ASM_IS`, `ASM_CF`, `ASM_BS`, `AIXA_IS`, etc.
5. Built `ASM_REPORTED` maps for the as-reported sheet (contract assets, ROU, evaluation tools, etc.)
6. Created `COMPANY_INFO` master dict tying everything together
7. Ran a verification function confirming all 400+ mapped rows match expected labels

**Key finding:** The CapIQ format is NOT identical across companies — row numbers differ by up to 20 rows between ASM and LRCX for CF/BS sections. Hardcoding row numbers from one company would have broken formulas for peers.

### Task 3: Model Skeleton + Input Tabs + Cover
**Subagent output:** DONE (125s)

Created `scripts/build_model.py` (250 lines) with:
- `build_cover(wb)` — Title page with company info, peer list, table of contents (17 sections)
- `copy_source_sheet(wb, source_wb, src_name, tgt_name)` — Generic function that copies all non-empty rows from source, applies blue font to numeric data, black font to labels
- `build_input_tabs(wb, source_wb)` — Creates 5 input tabs
- `main()` — Orchestrates build and saves to `output/ASM_Valuation_Model.xlsx`

**Output:** 6 sheets (Cover + 5 Input), zero errors

### Task 4: Standardized ASM
**Subagent output:** DONE_WITH_CONCERNS (799s) — Most complex task

Added `build_standardized_asm(wb)` (~390 lines) creating the "Std ASM" sheet with:

**Income Statement (rows 8-40):**
- Revenue → Cost of Sales → Gross Profit → SG&A → R&D → D&A → Other OpEx → Total OpEx → Recurring Operating Profit
- Non-recurring items: Gain/Loss on investments, Asset writedowns, Other unusual
- EBIT → Investment Income → Interest → FX → PBT → Tax → Net Income
- NOPAT = Recurring OP × (1 - ETR)
- EPS, DPS

**Balance Sheet Reformulated (rows 43-83):**
- Operating Working Capital: Trade Rec + Inventory + Prepaid - Payables - Accrued - Unearned Rev - Tax Payable - Other CL
- Net Non-Current Operating Assets: PP&E + Goodwill + Intangibles + Deferred Charges + Other - DTL - Other NCL
- Net Operating Assets = OWC + Net NCA
- Investment Assets (equity method + other)
- Business Assets = NOA + Investment Assets
- Financing: Debt + Cash = Net Debt, Equity, Invested Capital
- **BS Check = 0 for all 6 years**

**Cash Flow (rows 86-99):**
- NOPAT + D&A - Chg OWC - CapEx - Purch Intangibles = Oper CF
- + Net Inv Profit - Chg Inv Assets = FCF to D&E
- - Int after tax + Chg Debt = FCF to Equity

**Concerns raised & resolution:**
1. D&A double-counting in IS: Acknowledged as analytical presentation choice (D&A from CF is a memo line; COGS already includes it). Same approach as Sanofi/BioNTech examples.
2. BS used CapIQ items exclusively (not mixing with as-reported) to ensure balance. Correct decision.

### Task 5: Standardized Peers
**Subagent output:** DONE (668s)

Refactored `build_standardized_asm()` into generic `build_standardized_tab(wb, company_key, input_sheet, reported_sheet, output_sheet)`:
- Accepts any company key to look up correct row maps
- Handles missing map keys (writes 0 when a key doesn't exist in a peer's map)
- Peers without as-reported data: reported-specific rows (ROU, contract assets, etc.) default to 0
- Companies with financial debt (AMAT, LRCX) use CapIQ `total_debt`; lease-only companies (ASM, AIXA) sum lease liabilities
- Cash row uses `cash_st_investments` (includes ST investments)

Built 3 peer tabs: Std AIXTRON, Std AMAT, Std LRCX
**BS Check = 0 for all companies, all years**

### Task 6: Adjustments + Adjusted Statements
**Subagent output:** DONE (158s)

**Adjustments tab** — Text documentation of 6 areas examined:
1. R&D Capitalization: No adjustment (consistent with peers)
2. Equity Method Investment: Already separated in standardization
3. FX Gains/Losses: Already classified as non-operating
4. SBC: Already in operating expenses
5. Goodwill: No impairment needed
6. Asset Writedowns: Already non-recurring

**Adjusted tabs (4)** — Pass-through from Standardized tabs via cross-sheet references:
- Every data cell: `='Std ASM'!{col}{row}` (green font for cross-sheet ref convention)
- Labels copied as-is
- Structure preserved for downstream ratio/forecast references

This "adjustment overlay" structure means adjustments CAN be added later by modifying the Adj tab formulas without rebuilding anything.

### Task 7: Ratio Definitions + Ratio Analysis
**Subagent output:** DONE (220s)

**Ratio Definitions tab** — Text reference with formula descriptions, 5 sections:
- Profitability, Asset Management, Liquidity, Solvency, Growth

**Ratio tabs (4)** — `build_ratio_tab(wb, adj_sheet, output_sheet)`:
- Profitability: ROE, full DuPont decomposition (ROS × Asset Turnover × Leverage), RNOA (NOPAT Margin × NOA Turnover), Gross/EBIT/NOPAT margins
- Common-Size IS: Revenue=100%, COGS/SGA/RD as % of revenue
- Asset Management: Receivables/Inventory/Payables turnover & days, Cash Conversion Cycle, OWC/Revenue, PPE/Revenue
- Liquidity & Solvency: Cash/Debt, Debt-to-Equity, Interest Coverage
- Growth: Revenue Growth, Sustainable Growth
- Average-based ratios show N/A for 2019 (no prior year)
- All cells have division-by-zero guards

**Ratio Comparison tab** — 9 key ratios pulled from all 4 company ratio tabs side by side

**Verification:** ASM Gross Margin 2024 = 50.5% (1,481,373 / 2,932,724) ✓

### Task 8: Forecast Tab
**Subagent output:** DONE (281s)

Created `build_forecast_tab(wb)` with 5 sections:

**Forecast Assumptions (rows 5-16)** — Blue + yellow cells (user-editable):
| Assumption | 2025E | 2026E | 2027E | 2028E | 2029E | 2030E |
|-----------|-------|-------|-------|-------|-------|-------|
| Revenue Growth | 12% | 10% | 8% | 6% | 4% | 3% |
| Gross Margin | 50.5% | 50.5% | 51% | 51% | 51% | 51% |
| SG&A/Rev | 10.8% | 10.5% | 10.2% | 10% | 10% | 10% |
| R&D/Rev | 12.6% | 12.5% | 12.3% | 12% | 12% | 12% |
| D&A/Rev | 3.7% | 3.7% | 3.7% | 3.5% | 3.5% | 3.5% |
| Tax Rate | 21% | 21% | 21% | 21% | 21% | 21% |
| OWC/Rev | 5% | 5% | 5% | 5% | 5% | 5% |
| NCA/Rev | 40% | 40% | 39% | 38% | 37% | 36% |
| IA/Rev | 30% | 29% | 28% | 27% | 26% | 25% |
| Debt/Capital | 1% | 1% | 1% | 1% | 1% | 1% |
| Terminal Growth | | | | | | 2.5% |

Rationale: Revenue growth tapers from 12% (strong EUR 1.57bn backlog + AI/advanced packaging tailwinds) to 3% (GDP+ long-term). Margins match recent history with slight efficiency gains at scale. ASM is essentially debt-free (net cash position).

**Condensed IS, BS, CF** — Historical columns reference Adj ASM (green), forecast columns use formulas referencing assumptions (black). BS Check = 0 for all years.

### Tasks 9+10: Valuation Tab (combined)
**Subagent output:** DONE (305s)

Created `build_valuation_tab(wb)` (+1,089 lines) with 8 sections:

**1. Cost of Capital (WACC):**
- Risk-free rate: 2.5% (10yr German Bund)
- Beta: 1.30
- ERP: 5.0% (Damodaran)
- Cost of equity (CAPM): 9.0%
- Cost of debt: from actual interest/debt ratio
- WACC: weighted average using market-value weights

**2. DCF — FCF to Debt & Equity:**
- FCFF from Forecast tab (2025E-2030E)
- Gordon Growth terminal value: FCFF₂₀₃₀ × (1+g) / (WACC-g)
- Discount factors at WACC
- Enterprise Value = Sum of PV + PV of Terminal Value
- Equity Value = EV - Net Debt
- Implied Price = Equity Value / Shares × 1000

**3. DCF — FCF to Equity:**
- FCFE from Forecast tab
- Discounted at cost of equity
- Direct equity value (no net debt adjustment)

**4. PVAOI (Present Value of Abnormal Operating Income):**
- NOPAT from forecast
- Beginning BV of NOA
- Normal Earnings = WACC × Beginning NOA
- Abnormal NOPAT = NOPAT - Normal
- EV = Beginning NOA + Sum PV Abnormal + PV Terminal - Net Debt

**5. PVAE (Present Value of Abnormal Earnings):**
- Net Income from forecast
- Beginning BV of Equity
- Normal Earnings = Re × Beginning Equity
- Abnormal Earnings = NI - Normal
- Equity Value = Current BV Equity + Sum PV Abnormal

**6. Sensitivity Tables:**
- DCF FCFF grid: 11 WACC values (7%-12%) × 7 terminal growth values (1%-4%)
- PVAOI grid: Same dimensions
- Each cell contains a self-contained formula computing implied price for that WACC/g combination

**7. Comparable Company Analysis:**
- 3 peers with hardcoded EV, EBITDA, EBIT
- EV/EBITDA and EV/EBIT multiples computed
- Median multiples applied to ASM's 2024 metrics
- Implied price via EV bridge

**8. Football Field Summary:**
- All 6 methods listed with implied prices
- Average implied price
- BUY/SELL recommendation: `=IF(average > current_price, "BUY", "SELL")`

### Task 11: Final Polish
**Subagent output:** DONE (162s)

- Fixed freeze panes on Adjustments and Ratio Definitions tabs
- Adjusted freeze pane positions on Ratio and Valuation tabs
- Verified number format consistency across all 23 sheets
- Final build: 23 sheets, zero validation errors

---

## Phase 4: Final Output

### Workbook Structure (23 sheets)

| # | Sheet Name | Contents | Rows |
|---|-----------|----------|------|
| 1 | Cover | Title, company info, peers, ToC | ~35 |
| 2 | Input ASM | Raw CapIQ data (IS, CF, BS) | 230 |
| 3 | Input ASM (Reported) | As-reported financials | 128 |
| 4 | Input AIXTRON | Raw CapIQ peer data | 232 |
| 5 | Input AMAT | Raw CapIQ peer data | 241 |
| 6 | Input LRCX | Raw CapIQ peer data | 230 |
| 7 | Std ASM | Standardized IS/BS/CF | ~99 |
| 8 | Std AIXTRON | Standardized peer | ~99 |
| 9 | Std AMAT | Standardized peer | ~99 |
| 10 | Std LRCX | Standardized peer | ~99 |
| 11 | Adjustments | Accounting analysis documentation | ~60 |
| 12 | Adj ASM | Adjusted statements (pass-through) | ~99 |
| 13 | Adj AIXTRON | Adjusted peer | ~99 |
| 14 | Adj AMAT | Adjusted peer | ~99 |
| 15 | Adj LRCX | Adjusted peer | ~99 |
| 16 | Ratio Definitions | Formula reference | ~95 |
| 17 | Ratios ASM | Full ratio analysis | ~55 |
| 18 | Ratios AIXTRON | Peer ratios | ~55 |
| 19 | Ratios AMAT | Peer ratios | ~55 |
| 20 | Ratios LRCX | Peer ratios | ~55 |
| 21 | Ratio Comparison | Side-by-side peer comparison | ~50 |
| 22 | Forecast | Assumptions + IS/BS/CF projections | ~68 |
| 23 | Valuation | WACC, DCF, PVAOI, Comps, Sensitivity, Football Field | ~195 |

### Code Structure

| File | Lines | Purpose |
|------|-------|---------|
| `scripts/build_model.py` | ~3,200 | Main builder — all tab construction functions |
| `scripts/data_maps.py` | ~896 | Source data row/column mappings (verified) |
| `scripts/formatting.py` | ~198 | IB color coding and number format utilities |
| `scripts/recalc.py` | ~50 | Formula error validation |
| `scripts/__init__.py` | 0 | Package init |

### Git History (14 commits)

```
0503059 chore: update output workbook with final build
276d5b0 feat: complete ASM International valuation model with final polish
7a9948d feat: add Valuation tab with WACC, DCF, PVAOI, comps, sensitivity, football field
9dd9a33 feat: add Forecast tab with assumptions, IS, BS, and FCF projections
c9cade4 feat: add ratio definitions, ratio analysis, and peer comparison tabs
dcffd4f feat: add Adjustments documentation and Adjusted statement tabs
45db0c0 feat: add standardized tabs for AIXTRON, AMAT, LRCX peers
21d01b3 feat: add Standardized ASM tab with IS, BS, CF reformulation
f3d15ba feat: build model skeleton with Cover and Input tabs
4444f09 feat: add source data row/column mappings for all companies
3becf1d chore: add output/ directory with .gitkeep
0fbfc69 feat: add formatting utilities and validation script
81d95d6 Add implementation plan for ASM International valuation model
4b18e44 Add design spec for ASM International valuation model
156d5dc Initial commit
```

### Key Formatting Rules Applied

- **Blue text (0000FF):** All hardcoded input values (source data, assumptions)
- **Black text (000000):** All formulas and calculations
- **Green text (008000):** All cross-sheet references
- **Yellow background (FFFF00):** Key assumptions (forecast growth rates, WACC inputs, peer multiples)
- **Number formats:** Currency `#,##0;(#,##0);"-"`, Percentages `0.0%`, Multiples `0.0"x"`, Years as text `@`
- **Freeze panes:** Set on all data tabs so labels and year headers stay visible
- **Section headers:** Blue background fill with bold text
- **Total rows:** Top+bottom thin borders with bold text
- **Final totals:** Double-line bottom border

---

## Methodology Notes

### Why These Valuation Methods

The course requires both DCF and PVAOI. We implemented four methods for robustness:

1. **DCF (FCFF)** — Enterprise value approach. Discount free cash flows to all capital providers at WACC, subtract net debt. Most common in practice.

2. **DCF (FCFE)** — Equity value approach. Discount free cash flows available to equity holders at cost of equity. Should yield same result as FCFF if assumptions are consistent.

3. **PVAOI (Abnormal NOPAT)** — Accounting-based enterprise value. Starting from book value of net operating assets, adds present value of "abnormal" operating returns (NOPAT exceeding WACC × invested capital). Required by course spec.

4. **PVAE (Abnormal Earnings)** — Accounting-based equity value. Starting from book value of equity, adds present value of abnormal earnings (NI exceeding Re × equity). Complementary to PVAOI.

5. **Comparable Company Analysis** — Market-based. Apply peer multiples (EV/EBITDA, EV/EBIT) to ASM's metrics. Cross-check against intrinsic valuations.

### Why These Forecast Assumptions

- **Revenue Growth (12% → 3%):** ASM's 2024 backlog of EUR 1.57bn provides near-term visibility. AI chip demand and advanced packaging (Gate-All-Around transistors requiring ALD) drive growth. Growth tapers as the semiconductor cycle normalizes toward long-term GDP+ rates.

- **Gross Margin ~50.5-51%:** Consistent with ASM's 2022-2024 actuals. Slight improvement reflects mix shift toward higher-margin ALD products.

- **SG&A declining as % of revenue:** Operating leverage — fixed costs spread over growing revenue base.

- **R&D ~12-12.6%:** Semi equipment companies must invest heavily; ASM's ratio is consistent with peers.

- **Tax Rate 21%:** ASM's 2024 effective rate was 20.99%. Netherlands corporate rate is 25.8% but ASM benefits from the Innovation Box regime.

- **Debt/Capital 1%:** ASM is essentially debt-free with EUR 927M cash and only EUR 35M in lease liabilities.

### Data Integrity Checks

Every build run validates:
1. Zero Excel formula errors (no #REF!, #VALUE!, #DIV/0!, etc.)
2. Balance sheet check = 0 for all companies, all years
3. Revenue figures match source data exactly
4. Cross-sheet references resolve correctly

---

## How to Use

### Rebuild the Model
```bash
python3 scripts/build_model.py
```

### Validate
```bash
python3 scripts/recalc.py output/ASM_Valuation_Model.xlsx
```

### Modify Assumptions
Open `output/ASM_Valuation_Model.xlsx` in Excel. All blue+yellow cells on the Forecast and Valuation tabs are editable assumptions. Change any value and the entire model recalculates:
- Revenue growth rates → cascades through IS, BS, CF, all valuations
- WACC inputs (Rf, Beta, ERP) → cascades through all discount-based valuations
- Terminal growth rate → affects all terminal values
- Peer multiples → affects comparable analysis

### Regenerate with Different Data
Modify `ASM International and Comps data.xlsx` with updated CapIQ pulls, then re-run `python3 scripts/build_model.py`. The data_maps module handles the row mapping.
