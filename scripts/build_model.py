#!/usr/bin/env python3
"""Build the ASM International Valuation Model workbook."""

import sys
import os
import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from scripts.formatting import (
    BLUE_FONT,
    BLACK_FONT,
    FMT_CURRENCY,
    FMT_YEAR,
    style_input_cell,
    set_column_widths,
    freeze_panes,
)
from openpyxl.utils import get_column_letter
from scripts.data_maps import SOURCE_FILE

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
