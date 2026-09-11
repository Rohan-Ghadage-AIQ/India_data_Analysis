#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Batch Update for Top 100 Layers in Master Catalog and Child Sheets
===================================================================
Author: Senior Geospatial Analyst & Software Engineer

Execution Details:
1. Master Sheet ('GeoJSON Catalog'):
   - For rows 2 to 101 (Sr. No 1 to 100):
     * Updates 'How to Use This File' (Col 4) with deep 5-point Geospatial Analysis Protocols.
     * Updates 'Target Industries' (Col 5) with domain stakeholder listings.
     * Updates 'Source' (Col 6) with clean ArcGIS item URLs as clickable links.
     * Sets row height to 115 pt with text-wrapping.
   - Sets master column widths: A(10), B(42), C(55), D(58), E(35), F(50), G(20), H(20).
   - Updates AutoFilter: A1:H428.
2. Child Sheets (Sheets 1 to 100):
   - Updates Row 3 Table Headers to 6 columns:
     Col 1: 'Sr. No'
     Col 2: 'Layer Available Fields'
     Col 3: 'Target Industries'
     Col 4: 'Potential Analysis'
     Col 5: 'Potential Analysis Manual'  <-- KEPT VACANT FOR USER MANUAL ENTRY
     Col 6: 'Remark'
   - Updates all data rows (Row 4 to max_row):
     * Populates Col 3 ('Target Industries').
     * Populates Col 4 ('Potential Analysis') with Senior Geospatial Analyst depth.
     * Leaves Col 5 ('Potential Analysis Manual') strictly vacant (None/empty) with borders and zebra fill.
     * Populates Col 6 ('Remark') with clear, concise, purposeful context notes.
     * Sets row height to 48 pt.
   - Sets column widths: A(10), B(45), C(35), D(55), E(35), F(35).
   - Sets AutoFilter: A3:F{max_row}.
   - Preserves freeze panes (A4) and A1 '← Back to Catalog' hyperlinks.
"""

import os
import sys
import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

# Palette & Style Singletons
COLOR_PRIMARY_NAVY = "2F5496"
COLOR_PRIMARY_TEXT = "FFFFFF"
COLOR_LINK_BLUE = "0563C1"
COLOR_BORDER_LIGHT = "B4C6E7"
COLOR_BORDER_DARK = "1F3864"
COLOR_ZEBRA_EVEN = "F2F5F9"
COLOR_ZEBRA_ODD = "FFFFFF"
COLOR_BANNER_BG = "D9E1F2"

SIDE_THIN_LIGHT = Side(style="thin", color=COLOR_BORDER_LIGHT)
SIDE_THIN_DARK = Side(style="thin", color=COLOR_BORDER_DARK)
SIDE_MEDIUM_DARK = Side(style="medium", color=COLOR_BORDER_DARK)

BORDER_DATA_CELL = Border(
    left=SIDE_THIN_LIGHT, right=SIDE_THIN_LIGHT,
    top=SIDE_THIN_LIGHT, bottom=SIDE_THIN_LIGHT
)
BORDER_HEADER_CELL = Border(
    left=SIDE_THIN_DARK, right=SIDE_THIN_DARK,
    top=SIDE_THIN_DARK, bottom=SIDE_MEDIUM_DARK
)

FILL_HEADER = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
FILL_ZEBRA_EVEN = PatternFill(start_color=COLOR_ZEBRA_EVEN, end_color=COLOR_ZEBRA_EVEN, fill_type="solid")
FILL_ZEBRA_ODD = PatternFill(start_color=COLOR_ZEBRA_ODD, end_color=COLOR_ZEBRA_ODD, fill_type="solid")

FONT_HEADER = Font(name="Calibri", size=11, bold=True, color=COLOR_PRIMARY_TEXT)
FONT_DATA = Font(name="Calibri", size=10, color="000000")
FONT_LINK = Font(name="Calibri", size=10, color=COLOR_LINK_BLUE, underline="single")

ALIGN_CENTER_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT_TOP_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER_TOP = Alignment(horizontal="center", vertical="top")

# Path to pre-generated layer specifications
SPECS_JSON = r"C:\Users\RohanDhanajiGhadage\.gemini\antigravity-ide\brain\a4763728-6862-44d3-96c4-5d2e9c23804a\scratch\layer_specs_100.json"
TARGET_WORKBOOK = "exported_geojson_540andtif_catalog.xlsx"

# Import deep customized field logic for layers 1, 2, 3 if available
try:
    import update_catalog_and_sheets as custom_l123
    has_custom_l123 = True
except Exception:
    has_custom_l123 = False


def run_batch_update():
    print("=" * 80)
    print("BATCH UPDATE FOR TOP 100 LAYERS (MASTER CATALOG & CHILD SHEETS)")
    print("=" * 80)
    
    if not os.path.exists(TARGET_WORKBOOK):
        print(f"[!] Error: Target workbook '{TARGET_WORKBOOK}' not found.")
        sys.exit(1)
        
    if not os.path.exists(SPECS_JSON):
        print(f"[!] Error: Specifications file '{SPECS_JSON}' not found.")
        sys.exit(1)
        
    with open(SPECS_JSON, "r", encoding="utf-8") as f:
        specs = json.load(f)
        
    print(f"[*] Loaded specifications for {len(specs)} layers.")
    print(f"[*] Loading workbook: '{TARGET_WORKBOOK}' ...")
    wb = openpyxl.load_workbook(TARGET_WORKBOOK)
    
    # -------------------------------------------------------------------------
    # PART 1: UPDATE MASTER CATALOG ('GeoJSON Catalog')
    # -------------------------------------------------------------------------
    print("\n" + "=" * 50)
    print("[1/2] Updating Master Sheet: 'GeoJSON Catalog'")
    print("=" * 50)
    
    cat_ws = wb["GeoJSON Catalog"]
    
    # Identify column indices
    col_map = {}
    for c in range(1, cat_ws.max_column + 1):
        v = str(cat_ws.cell(1, c).value or "").strip()
        if v:
            col_map[v] = c
            
    print(f"    Master Columns: {col_map}")
    how_col = col_map.get("How to Use This File", 4)
    ind_col = col_map.get("Target Industries", 5)
    src_col = col_map.get("Source", 6)
    
    updated_master = 0
    for sr in range(1, 101):
        row_idx = sr + 1
        spec = specs.get(str(sr))
        if not spec:
            continue
            
        # 1. Update How to Use This File (Column 4)
        c_how = cat_ws.cell(row=row_idx, column=how_col, value=spec['master_how_to_use'])
        c_how.font = FONT_DATA
        c_how.alignment = ALIGN_LEFT_TOP_WRAP
        cat_ws.row_dimensions[row_idx].height = 115
        
        # 2. Update Target Industries (Column 5)
        c_ind = cat_ws.cell(row=row_idx, column=ind_col, value=spec['master_target_industries'])
        c_ind.font = FONT_DATA
        c_ind.alignment = ALIGN_LEFT_TOP_WRAP
        
        # 3. Update Source URL (Column 6)
        url = spec.get('url', '')
        if url:
            c_src = cat_ws.cell(row=row_idx, column=src_col, value=url)
            c_src.hyperlink = url
            c_src.font = FONT_LINK
            c_src.alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)
            
        updated_master += 1
        
    print(f"[+] Successfully updated {updated_master} rows in 'GeoJSON Catalog'.")
    
    # Master column widths
    master_widths = {
        "A": 10, "B": 42, "C": 55, "D": 58,
        "E": 35, "F": 50, "G": 20, "H": 20
    }
    for col_letter, w in master_widths.items():
        cat_ws.column_dimensions[col_letter].width = w
        
    cat_ws.auto_filter.ref = f"A1:H{cat_ws.max_row}"
    print(f"[+] Master AutoFilter updated to: {cat_ws.auto_filter.ref}")

    # -------------------------------------------------------------------------
    # PART 2: UPDATE CHILD SHEETS (SHEETS 1 TO 100)
    # -------------------------------------------------------------------------
    print("\n" + "=" * 50)
    print("[2/2] Updating Child Sheets (1 to 100)")
    print("=" * 50)
    
    child_headers = [
        ("Sr. No", 10),
        ("Layer Available Fields", 45),
        ("Target Industries", 35),
        ("Potential Analysis", 55),
        ("Potential Analysis Manual", 35),  # Vacant column for user
        ("Remark", 35),
    ]
    
    updated_sheets_count = 0
    total_fields_updated = 0
    
    for sr in range(1, 101):
        spec = specs.get(str(sr))
        if not spec:
            continue
            
        sheet_name = spec['sheet_name']
        if sheet_name not in wb.sheetnames:
            print(f"[!] Warning: Sheet '{sheet_name}' (Sr {sr}) not found in workbook.")
            continue
            
        ws = wb[sheet_name]
        
        # 1. Update Row 3 Table Headers
        ws.row_dimensions[3].height = 26
        for col_idx, (hdr_text, col_w) in enumerate(child_headers, start=1):
            cell = ws.cell(row=3, column=col_idx, value=hdr_text)
            cell.font = FONT_HEADER
            cell.fill = FILL_HEADER
            cell.alignment = ALIGN_CENTER_CENTER
            cell.border = BORDER_HEADER_CELL
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = col_w
            
        # 2. Update Data Rows (Row 4 onwards)
        field_dict = spec.get('field_data', {})
        max_r = ws.max_row
        
        for r in range(4, max_r + 1):
            f_name = str(ws.cell(r, 2).value or "").strip()
            if not f_name:
                continue
                
            fill_style = FILL_ZEBRA_ODD if (r % 2 == 0) else FILL_ZEBRA_EVEN
            
            # Use specialized logic for layers 1, 2, 3 if available, else spec
            if sr == 1 and has_custom_l123:
                t_ind, p_an, rem = custom_l123.get_layer1_field_info(f_name)
            elif sr == 2 and has_custom_l123:
                t_ind, p_an, rem = custom_l123.get_layer2_field_info(f_name)
            elif sr == 3 and has_custom_l123:
                t_ind, p_an, rem = custom_l123.get_layer3_field_info(f_name)
            else:
                f_info = field_dict.get(f_name, {})
                t_ind = f_info.get('target_industries', spec['master_target_industries'])
                p_an = f_info.get('potential_analysis', f"Spatial distribution analysis and infrastructure overlay for {f_name}")
                rem = f_info.get('remark', "")
                
            # Col 1: Sr. No
            c1 = ws.cell(row=r, column=1, value=r - 3)
            c1.font = FONT_DATA
            c1.alignment = ALIGN_CENTER_TOP
            c1.border = BORDER_DATA_CELL
            c1.fill = fill_style
            
            # Col 2: Field Name
            c2 = ws.cell(row=r, column=2, value=f_name)
            c2.font = FONT_DATA
            c2.alignment = ALIGN_LEFT_TOP_WRAP
            c2.border = BORDER_DATA_CELL
            c2.fill = fill_style
            
            # Col 3: Target Industries
            c3 = ws.cell(row=r, column=3, value=t_ind)
            c3.font = FONT_DATA
            c3.alignment = ALIGN_LEFT_TOP_WRAP
            c3.border = BORDER_DATA_CELL
            c3.fill = fill_style
            
            # Col 4: Potential Analysis (Senior Geospatial Analyst Depth)
            c4 = ws.cell(row=r, column=4, value=p_an)
            c4.font = FONT_DATA
            c4.alignment = ALIGN_LEFT_TOP_WRAP
            c4.border = BORDER_DATA_CELL
            c4.fill = fill_style
            
            # Col 5: Potential Analysis Manual (STRICTLY VACANT FOR USER MANUAL ENTRY)
            c5 = ws.cell(row=r, column=5)
            c5.value = None
            c5.font = FONT_DATA
            c5.alignment = ALIGN_LEFT_TOP_WRAP
            c5.border = BORDER_DATA_CELL
            c5.fill = fill_style
            
            # Col 6: Remark (Clear, concise context note)
            c6 = ws.cell(row=r, column=6, value=rem)
            c6.font = FONT_DATA
            c6.alignment = ALIGN_LEFT_TOP_WRAP
            c6.border = BORDER_DATA_CELL
            c6.fill = fill_style
            
            ws.row_dimensions[r].height = 48
            total_fields_updated += 1
            
        # Update AutoFilter on child sheet to cover 6 columns
        ws.auto_filter.ref = f"A3:F{max_r}"
        
        # Ensure freeze panes remain at A4
        ws.freeze_panes = "A4"
        
        # Ensure A1 hyperlink back to catalog row
        back_cell = ws.cell(row=1, column=1)
        back_cell.value = "← Back to Catalog"
        back_cell.font = FONT_LINK
        back_cell.fill = PatternFill(start_color=COLOR_BANNER_BG, end_color=COLOR_BANNER_BG, fill_type="solid")
        back_cell.alignment = ALIGN_CENTER_CENTER
        back_cell.border = BORDER_DATA_CELL
        back_cell.hyperlink = f"#'GeoJSON Catalog'!A{sr + 1}"
        
        updated_sheets_count += 1
        if updated_sheets_count % 10 == 0 or updated_sheets_count == 100:
            print(f"    -> Processed {updated_sheets_count}/100 child sheets ({total_fields_updated} fields updated)...")

    # -------------------------------------------------------------------------
    # PART 3: SAVE WORKBOOK
    # -------------------------------------------------------------------------
    print("\n" + "=" * 50)
    print(f"[*] Saving updated workbook to '{TARGET_WORKBOOK}' ...")
    wb.save(TARGET_WORKBOOK)
    print(f"[+] Successfully saved '{TARGET_WORKBOOK}'!")
    print("=" * 80)
    print("BATCH UPDATE COMPLETE:")
    print(f"  - Master Rows Updated:  {updated_master} (Rows 2 to 101)")
    print(f"  - Child Sheets Updated: {updated_sheets_count} (Sheets 1 to 100)")
    print(f"  - Total Fields Processed: {total_fields_updated}")
    print(f"  - 'Potential Analysis Manual' Column Added & Kept Vacant Across All 100 Sheets")
    print("=" * 80)


if __name__ == "__main__":
    run_batch_update()
