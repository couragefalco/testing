#!/usr/bin/env python3
"""
Validation script that scans an openpyxl workbook for Excel error tokens.

Usage:
    python3 scripts/recalc.py [path_to_workbook]

Defaults to output/ASM_Valuation_Model.xlsx when no argument is given.
Returns exit code 0 if clean, 1 if errors are found.
"""

import sys
import os

import openpyxl

# Excel error tokens to scan for
ERROR_TOKENS = {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#NULL!", "#N/A"}


def scan_workbook(path):
    """Open *path* and return a list of ``(sheet!cell, value)`` tuples for
    every cell whose value is an Excel error token."""
    if not os.path.isfile(path):
        print(f"ERROR: File not found: {path}")
        sys.exit(2)

    wb = openpyxl.load_workbook(path, data_only=True)
    errors = []

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip() in ERROR_TOKENS:
                    ref = f"{ws.title}!{cell.coordinate}"
                    errors.append((ref, cell.value.strip()))

    wb.close()
    return errors


def main():
    default_path = os.path.join("output", "ASM_Valuation_Model.xlsx")
    path = sys.argv[1] if len(sys.argv) > 1 else default_path

    print(f"Scanning: {path}")
    errors = scan_workbook(path)

    if errors:
        print(f"\n{len(errors)} error(s) found:\n")
        for ref, value in errors:
            print(f"  {ref}  ->  {value}")
        sys.exit(1)
    else:
        print("No errors found.")
        sys.exit(0)


if __name__ == "__main__":
    main()
