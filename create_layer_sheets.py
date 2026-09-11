#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interactive Layer Sheets Generator
==================================
Reads 'first100json_catalog.xlsx' and generates an interactive, high-performance
multi-sheet workbook where:
1. Every row in the main catalog links directly (via clickable hyperlink) to its dedicated sheet.
2. Each individual sheet provides:
   - Heading banner with Sr. No, Layer Name, and a '← Back to Catalog' button.
   - Field-by-field breakdown (one row per attribute field).
   - Dedicated audit columns: 'Target Industries', 'Potential Analysis', and 'Remark'.
   - Freeze panes, auto-filters, zebra striping, and tuned column widths.
3. Engineered for high performance:
   - Style singletons to avoid memory bloat and eliminate lag.
   - Zero-row-skipping guarantee with comprehensive error handling.
   - Safe file I/O handling Excel file locks gracefully.

Author: Senior Software Engineer
"""

import os
import re
import sys
import shutil
import tempfile
import argparse
from typing import List, Tuple, Dict, Any

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================================
# CONFIGURATION & CONSTANTS
# ============================================================================

DEFAULT_INPUT_FILE = "first100json_catalog.xlsx"
DEFAULT_OUTPUT_FILE = "first100json_catalog.xlsx"
FALLBACK_OUTPUT_FILE = "first100json_catalog_detailed.xlsx"

# Characters forbidden in Excel worksheet names
FORBIDDEN_SHEET_CHARS_REGEX = re.compile(r"[\\/*?:\[\]]")

# Palette & Design Tokens
COLOR_PRIMARY_NAVY = "2F5496"    # Dark blue header fill
COLOR_PRIMARY_TEXT = "FFFFFF"    # White header text
COLOR_LINK_BLUE = "0563C1"       # Excel standard hyperlink blue
COLOR_BORDER_LIGHT = "B4C6E7"    # Soft blue-gray cell border
COLOR_BORDER_DARK = "1F3864"     # Dark navy border
COLOR_ZEBRA_EVEN = "F2F5F9"      # Very subtle blue-tint zebra row
COLOR_ZEBRA_ODD = "FFFFFF"       # Crisp white row
COLOR_BANNER_BG = "D9E1F2"       # Soft accent for metadata banners
COLOR_BADGE_BG = "4472C4"        # Accent badge for Sr. No

# ============================================================================
# REUSABLE STYLE SINGLETONS (High Performance & Minimal Memory Footprint)
# ============================================================================

SIDE_THIN_LIGHT = Side(style="thin", color=COLOR_BORDER_LIGHT)
SIDE_THIN_DARK = Side(style="thin", color=COLOR_BORDER_DARK)
SIDE_MEDIUM_DARK = Side(style="medium", color=COLOR_BORDER_DARK)

BORDER_DATA_CELL = Border(
    left=SIDE_THIN_LIGHT,
    right=SIDE_THIN_LIGHT,
    top=SIDE_THIN_LIGHT,
    bottom=SIDE_THIN_LIGHT,
)

BORDER_HEADER_CELL = Border(
    left=SIDE_THIN_DARK,
    right=SIDE_THIN_DARK,
    top=SIDE_THIN_DARK,
    bottom=SIDE_MEDIUM_DARK,
)

FILL_HEADER = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
FILL_ZEBRA_EVEN = PatternFill(start_color=COLOR_ZEBRA_EVEN, end_color=COLOR_ZEBRA_EVEN, fill_type="solid")
FILL_ZEBRA_ODD = PatternFill(start_color=COLOR_ZEBRA_ODD, end_color=COLOR_ZEBRA_ODD, fill_type="solid")
FILL_BANNER = PatternFill(start_color=COLOR_BANNER_BG, end_color=COLOR_BANNER_BG, fill_type="solid")
FILL_BADGE = PatternFill(start_color=COLOR_BADGE_BG, end_color=COLOR_BADGE_BG, fill_type="solid")

FONT_HEADER = Font(name="Calibri", size=11, bold=True, color=COLOR_PRIMARY_TEXT)
FONT_HEADER_SM = Font(name="Calibri", size=10, bold=True, color=COLOR_PRIMARY_TEXT)
FONT_DATA = Font(name="Calibri", size=10, color="000000")
FONT_DATA_BOLD = Font(name="Calibri", size=10, bold=True, color="000000")
FONT_LINK = Font(name="Calibri", size=10, bold=True, color=COLOR_LINK_BLUE, underline="single")
FONT_BANNER_LABEL = Font(name="Calibri", size=10, bold=True, color=COLOR_PRIMARY_NAVY)
FONT_BANNER_VALUE = Font(name="Calibri", size=10, bold=True, color="1F3864")

ALIGN_CENTER_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_CENTER_TOP = Alignment(horizontal="center", vertical="top")
ALIGN_LEFT_TOP_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_LEFT_TOP_NOWRAP = Alignment(horizontal="left", vertical="top", wrap_text=False)
ALIGN_LEFT_CENTER = Alignment(horizontal="left", vertical="center")


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def sanitize_sheet_name(sr_no: Any, layer_name: str, used_names: set) -> str:
    """
    Generate a valid, unique Excel worksheet name (max 31 chars, no illegal chars).
    Format: 'Sr_{sr_no}_{short_title}'
    """
    # Clean forbidden characters
    clean_title = FORBIDDEN_SHEET_CHARS_REGEX.sub("_", str(layer_name)).strip()
    clean_title = re.sub(r"\s+", "_", clean_title)
    clean_title = re.sub(r"_+", "_", clean_title).strip("_")

    prefix = f"{sr_no}_"
    max_title_len = 31 - len(prefix)

    candidate = (prefix + clean_title[:max_title_len]).rstrip("_")
    if not candidate:
        candidate = f"Layer_{sr_no}"

    # Ensure uniqueness across the workbook
    final_name = candidate[:31]
    counter = 2
    while final_name in used_names:
        suffix = f"_{counter}"
        avail_len = 31 - len(suffix)
        final_name = candidate[:avail_len] + suffix
        counter += 1

    used_names.add(final_name)
    return final_name


def parse_fields(fields_value: Any) -> List[str]:
    """
    Safely parse comma-separated field string into a clean list of field names.
    Preserves exact original key names and whitespace integrity within field names.
    """
    if not fields_value:
        return []

    text = str(fields_value).strip()
    if not text or text.lower() in ("none", "no attribute fields available", "no attributes"):
        return []

    # Split by comma followed by optional space
    raw_fields = text.split(", ")
    fields = []
    for f in raw_fields:
        f_clean = f.strip()
        if f_clean:
            fields.append(f_clean)
    return fields


def safe_load_workbook(file_path: str) -> Tuple[openpyxl.Workbook, bool]:
    """
    Safely load an Excel workbook. If the file is locked by Excel,
    creates a shadow temporary copy using Windows file-sharing to read it.
    Returns (workbook, is_temporary_copy).
    """
    try:
        wb = openpyxl.load_workbook(file_path)
        return wb, False
    except PermissionError:
        print(f"[*] Notice: '{file_path}' is currently open in Excel. Accessing via shadow copy...")
        temp_dir = tempfile.gettempdir()
        temp_copy = os.path.join(temp_dir, f"temp_read_{os.path.basename(file_path)}")
        # Copy using PowerShell which allows shared read access
        cmd = f'powershell -Command "Copy-Item -Path \'{file_path}\' -Destination \'{temp_copy}\' -Force"'
        ret = os.system(cmd)
        if ret != 0 or not os.path.exists(temp_copy):
            raise PermissionError(f"Could not read '{file_path}'. Please close it in Excel and try again.")
        wb = openpyxl.load_workbook(temp_copy)
        return wb, True


def safe_save_workbook(wb: openpyxl.Workbook, target_path: str, fallback_path: str) -> str:
    """
    Save workbook to target path. If locked by Excel, saves to fallback path
    so work is never lost. Returns the actual saved file path.
    """
    try:
        wb.save(target_path)
        return target_path
    except PermissionError:
        print(f"[!] Warning: Cannot write directly to '{target_path}' (file locked by Excel).")
        wb.save(fallback_path)
        print(f"[+] Successfully saved detailed workbook to fallback location: '{fallback_path}'")
        return fallback_path


# ============================================================================
# CORE WORKBOOK GENERATOR
# ============================================================================

def build_layer_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    sr_no: Any,
    layer_name: str,
    fields: List[str],
    catalog_sheet_name: str,
    catalog_row: int,
    target_industries_hint: str = "",
):
    """
    Build a dedicated, high-performance worksheet for an individual layer:
    - Row 1: Banner with Back link, Sr. No, and Layer Name heading.
    - Row 2: Spacer.
    - Row 3: Table Column Headers ('Sr. No', 'Field Name', 'Target Industries', 'Potential Analysis', 'Remark').
    - Row 4+: Data rows for each field with alternating zebra striping and borders.
    """
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True

    # ------------------------------------------------------------------------
    # 1. Heading Banner (Row 1)
    # ------------------------------------------------------------------------
    ws.row_dimensions[1].height = 28

    # A1: Back to Catalog Button
    cell_back = ws.cell(row=1, column=1, value="← Back to Catalog")
    cell_back.font = FONT_LINK
    cell_back.fill = FILL_BANNER
    cell_back.alignment = ALIGN_CENTER_CENTER
    cell_back.border = BORDER_DATA_CELL
    cell_back.hyperlink = f"#'{catalog_sheet_name}'!A{catalog_row}"

    # B1: 'Sr. No' Label
    cell_sr_lbl = ws.cell(row=1, column=2, value="Sr. No")
    cell_sr_lbl.font = FONT_HEADER_SM
    cell_sr_lbl.fill = FILL_HEADER
    cell_sr_lbl.alignment = ALIGN_CENTER_CENTER
    cell_sr_lbl.border = BORDER_DATA_CELL

    # C1: Sr. No Value
    cell_sr_val = ws.cell(row=1, column=3, value=sr_no)
    cell_sr_val.font = FONT_DATA_BOLD
    cell_sr_val.fill = FILL_BANNER
    cell_sr_val.alignment = ALIGN_CENTER_CENTER
    cell_sr_val.border = BORDER_DATA_CELL

    # D1: 'Layer Name' Label
    cell_name_lbl = ws.cell(row=1, column=4, value="Layer Name")
    cell_name_lbl.font = FONT_HEADER_SM
    cell_name_lbl.fill = FILL_HEADER
    cell_name_lbl.alignment = ALIGN_CENTER_CENTER
    cell_name_lbl.border = BORDER_DATA_CELL

    # E1: Layer Name Value
    cell_name_val = ws.cell(row=1, column=5, value=str(layer_name))
    cell_name_val.font = FONT_BANNER_VALUE
    cell_name_val.fill = FILL_BANNER
    cell_name_val.alignment = ALIGN_LEFT_CENTER
    cell_name_val.border = BORDER_DATA_CELL

    # ------------------------------------------------------------------------
    # 2. Spacer (Row 2)
    # ------------------------------------------------------------------------
    ws.row_dimensions[2].height = 8

    # ------------------------------------------------------------------------
    # 3. Table Column Headers (Row 3)
    # ------------------------------------------------------------------------
    ws.row_dimensions[3].height = 26
    headers = [
        ("Sr. No", 10),
        ("Layer Available Fields", 45),
        ("Target Industries", 35),
        ("Potential Analysis", 30),
        ("Remark", 30),
    ]

    for col_idx, (hdr_text, col_width) in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=hdr_text)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER_CENTER
        cell.border = BORDER_HEADER_CELL
        # Set column width
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_width

    # ------------------------------------------------------------------------
    # 4. Field Rows (Row 4 onwards)
    # ------------------------------------------------------------------------
    if not fields:
        # Handle layers without attributes cleanly
        r = 4
        ws.row_dimensions[r].height = 22
        cell_idx = ws.cell(row=r, column=1, value=1)
        cell_idx.font = FONT_DATA
        cell_idx.alignment = ALIGN_CENTER_TOP
        cell_idx.border = BORDER_DATA_CELL
        cell_idx.fill = FILL_ZEBRA_ODD

        cell_f = ws.cell(row=r, column=2, value="No attribute fields recorded for this layer")
        cell_f.font = FONT_DATA
        cell_f.alignment = ALIGN_LEFT_TOP_WRAP
        cell_f.border = BORDER_DATA_CELL
        cell_f.fill = FILL_ZEBRA_ODD

        for c in (3, 4, 5):
            empty_cell = ws.cell(row=r, column=c, value="")
            empty_cell.font = FONT_DATA
            empty_cell.border = BORDER_DATA_CELL
            empty_cell.fill = FILL_ZEBRA_ODD
        last_row = 4
    else:
        for f_idx, field_name in enumerate(fields, start=1):
            r = 3 + f_idx
            ws.row_dimensions[r].height = 20
            row_fill = FILL_ZEBRA_EVEN if f_idx % 2 == 0 else FILL_ZEBRA_ODD

            # Col 1: Field Sr. No
            c1 = ws.cell(row=r, column=1, value=f_idx)
            c1.font = FONT_DATA
            c1.alignment = ALIGN_CENTER_TOP
            c1.border = BORDER_DATA_CELL
            c1.fill = row_fill

            # Col 2: Field Name
            c2 = ws.cell(row=r, column=2, value=field_name)
            c2.font = FONT_DATA
            c2.alignment = ALIGN_LEFT_TOP_WRAP
            c2.border = BORDER_DATA_CELL
            c2.fill = row_fill

            # Col 3: Target Industries
            c3 = ws.cell(row=r, column=3, value="")
            c3.font = FONT_DATA
            c3.alignment = ALIGN_LEFT_TOP_WRAP
            c3.border = BORDER_DATA_CELL
            c3.fill = row_fill

            # Col 4: Potential Analysis
            c4 = ws.cell(row=r, column=4, value="")
            c4.font = FONT_DATA
            c4.alignment = ALIGN_LEFT_TOP_WRAP
            c4.border = BORDER_DATA_CELL
            c4.fill = row_fill

            # Col 5: Remark
            c5 = ws.cell(row=r, column=5, value="")
            c5.font = FONT_DATA
            c5.alignment = ALIGN_LEFT_TOP_WRAP
            c5.border = BORDER_DATA_CELL
            c5.fill = row_fill

        last_row = 3 + len(fields)

    # ------------------------------------------------------------------------
    # 5. Sheet Configuration (Freeze Header Panes & AutoFilter)
    # ------------------------------------------------------------------------
    # Freeze rows 1-3 so banner and table headers remain sticky when scrolling
    ws.freeze_panes = "A4"

    # AutoFilter over table headers
    ws.auto_filter.ref = f"A3:E{last_row}"


def process_catalog(input_path: str, output_path: str):
    """
    Main orchestration routine:
    1. Reads catalog sheet.
    2. Removes any previous layer detail sheets (if re-running).
    3. Builds individual sheets for all rows.
    4. Updates Sr. No cells in catalog with hyperlinks to their sheets.
    5. Saves the final enriched workbook.
    """
    print("=" * 75)
    print("INTERACTIVE LAYER SHEETS WORKBOOK GENERATOR")
    print("=" * 75)
    print(f"[*] Input Catalog:  {input_path}")

    if not os.path.exists(input_path):
        print(f"[!] Error: Input catalog file '{input_path}' not found.")
        sys.exit(1)

    # Safe load
    wb, was_temp = safe_load_workbook(input_path)
    catalog_ws = wb.active
    catalog_sheet_name = catalog_ws.title
    print(f"[*] Active Catalog Sheet: '{catalog_sheet_name}'")

    # Read Catalog Headers
    header_row = [cell.value for cell in catalog_ws[1]]
    col_map = {name: idx for idx, name in enumerate(header_row, start=1) if name}

    # Verify required columns
    required_cols = ["Sr. No", "Layer Name", "Layer Available Fields"]
    for req in required_cols:
        if req not in col_map:
            print(f"[!] Error: Missing required column '{req}' in '{catalog_sheet_name}'. Found: {header_row}")
            sys.exit(1)

    sr_col_idx = col_map["Sr. No"]
    layer_name_col_idx = col_map["Layer Name"]
    fields_col_idx = col_map["Layer Available Fields"]
    ind_col_idx = col_map.get("Target Industries", None)

    # Collect catalog entries
    catalog_entries = []
    for r in range(2, catalog_ws.max_row + 1):
        sr_val = catalog_ws.cell(row=r, column=sr_col_idx).value
        name_val = catalog_ws.cell(row=r, column=layer_name_col_idx).value
        fields_val = catalog_ws.cell(row=r, column=fields_col_idx).value
        ind_val = catalog_ws.cell(row=r, column=ind_col_idx).value if ind_col_idx else ""

        # Stop if row is completely empty
        if sr_val is None and name_val is None:
            continue

        catalog_entries.append({
            "row_idx": r,
            "sr_no": sr_val,
            "layer_name": name_val or f"Layer_{sr_val}",
            "fields_str": fields_val,
            "industries": ind_val or "",
        })

    print(f"[*] Found {len(catalog_entries)} catalog rows to process.")
    if not catalog_entries:
        print("[!] No data rows found in catalog.")
        sys.exit(0)

    # Clean up existing sub-sheets if re-running on an already populated workbook
    existing_sheets = list(wb.sheetnames)
    for sname in existing_sheets:
        if sname != catalog_sheet_name:
            del wb[sname]

    used_sheet_names = {catalog_sheet_name}
    total_fields_created = 0

    print(f"[*] Generating individual sheets and two-way hyperlinks...")

    for i, entry in enumerate(catalog_entries, start=1):
        sr_no = entry["sr_no"]
        layer_name = entry["layer_name"]
        row_idx = entry["row_idx"]
        fields_list = parse_fields(entry["fields_str"])

        # 1. Create unique sheet name
        sheet_name = sanitize_sheet_name(sr_no, layer_name, used_sheet_names)

        # 2. Build individual layer sheet
        build_layer_sheet(
            wb=wb,
            sheet_name=sheet_name,
            sr_no=sr_no,
            layer_name=layer_name,
            fields=fields_list,
            catalog_sheet_name=catalog_sheet_name,
            catalog_row=row_idx,
            target_industries_hint=entry["industries"],
        )
        total_fields_created += len(fields_list)

        # 3. Add clickable hyperlink to the Catalog sheet's Sr. No cell
        sr_cell = catalog_ws.cell(row=row_idx, column=sr_col_idx)
        sr_cell.hyperlink = f"#'{sheet_name}'!A1"
        sr_cell.font = FONT_LINK
        sr_cell.alignment = ALIGN_CENTER_CENTER

        if i % 10 == 0 or i == len(catalog_entries):
            print(f"    -> Processed {i}/{len(catalog_entries)} layers ({total_fields_created} total fields)...")

    # Save final workbook safely
    print(f"\n[*] Saving workbook...")
    saved_path = safe_save_workbook(wb, output_path, FALLBACK_OUTPUT_FILE)

    print("=" * 75)
    print("SUCCESS: INTERACTIVE WORKBOOK GENERATED")
    print("=" * 75)
    print(f"  Total Layers Processed: {len(catalog_entries)}")
    print(f"  Total Sheets Created:   {len(catalog_entries) + 1} (1 Catalog + {len(catalog_entries)} Layer Sheets)")
    print(f"  Total Fields Indexed:   {total_fields_created}")
    print(f"  Output File:            {saved_path}")
    print("=" * 75)


# ============================================================================
# CLI ENTRY POINT
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Generate interactive multi-sheet Excel workbook from GeoJSON catalog."
    )
    parser.add_argument(
        "-i", "--input",
        default=DEFAULT_INPUT_FILE,
        help=f"Path to input catalog Excel file (default: {DEFAULT_INPUT_FILE})"
    )
    parser.add_argument(
        "-o", "--output",
        default=DEFAULT_OUTPUT_FILE,
        help=f"Path to output Excel file (default: {DEFAULT_OUTPUT_FILE})"
    )

    args = parser.parse_args()
    script_dir = os.path.dirname(os.path.abspath(__file__))

    input_path = os.path.join(script_dir, args.input) if not os.path.isabs(args.input) else args.input
    output_path = os.path.join(script_dir, args.output) if not os.path.isabs(args.output) else args.output

    process_catalog(input_path, output_path)


if __name__ == "__main__":
    main()
