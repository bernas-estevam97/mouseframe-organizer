import os
from datetime import datetime
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from .parser import find_table_data, group_data_by_tag, sort_tags
from .mapping import load_mapping_file, match_table_id

# Visual Representation Modes
MODE_PASTEL_ACCENTS = "Pastel Header Tag + Explicit Columns"
MODE_COLUMNS_ONLY = "Explicit Columns Only (Clean White Grid)"
MODE_SOFT_ROW_FILLS = "Soft Pastel Row Fills"

VISUAL_MODES = [MODE_PASTEL_ACCENTS, MODE_COLUMNS_ONLY, MODE_SOFT_ROW_FILLS]

# High-contrast, clean pastel palettes (Hex, Font Color Hex)
MALE_PASTELS = [
    ("EBF3FA", "1B365D"),  # Soft Sky Blue
    ("E3F2FD", "0D47A1"),  # Light Blue
    ("E0F7FA", "006064"),  # Cyan Tint
    ("F0F4C3", "33691E"),  # Soft Lime
]

FEMALE_PASTELS = [
    ("FDE8E8", "7A1C1C"),  # Soft Rose
    ("FCE4EC", "880E4F"),  # Light Pink
    ("FFF3E0", "E65100"),  # Soft Peach
    ("F3E5F5", "4A148C"),  # Soft Lavender
]

NEUTRAL_PASTELS = [
    ("F5F5F5", "212121"),  # Light Gray
    ("FFFDE7", "F57F17"),  # Light Cream
]


def _get_pastel_fill(genotype, sex, genotype_index):
    """Generates soft pastel PatternFill and Font for group visualization."""
    g = str(genotype).strip().lower()
    s = str(sex).strip().lower()

    if g not in genotype_index:
        return None, None

    idx = genotype_index[g]

    if s in ("m", "male", "macho"):
        palette = MALE_PASTELS
    elif s in ("f", "female", "fêmica", "fêmea"):
        palette = FEMALE_PASTELS
    else:
        palette = NEUTRAL_PASTELS

    bg_hex, font_hex = palette[idx % len(palette)]
    return PatternFill(start_color="FF" + bg_hex, end_color="FF" + bg_hex, fill_type="solid"), Font(color="FF" + font_hex, bold=True)


def process_and_export_file(
    input_file,
    mapping_file=None,
    output_dir=None,
    threshold=500.0,
    visual_mode=MODE_PASTEL_ACCENTS,
    generate_summary=True,
    progress_callback=None
):
    """
    Main pipeline function to extract raw Mouseframe data, map metadata,
    apply visual grouping, generate summary statistics per animal/timepoint, and export Excel.
    """
    if not os.path.isfile(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    if progress_callback:
        progress_callback("Extracting table data from Excel workbook...", 10)

    data = find_table_data(input_file, threshold=threshold)
    if not data:
        raise ValueError("No valid 'Table ID' entries found in the input file.")

    if progress_callback:
        progress_callback(f"Extracted {len(data)} tables. Loading mapping file...", 30)

    mapping_dict, genotype_index, unique_genotypes = load_mapping_file(mapping_file)

    w_groups = group_data_by_tag(data)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    base_dir = output_dir if (output_dir and os.path.isdir(output_dir)) else os.path.dirname(input_file)
    output_file = os.path.join(base_dir, f"MFrame_clean_output_{timestamp}.xlsx")

    wb = Workbook()
    wb.remove(wb.active)  # Remove default active sheet

    # Define common Excel styles
    header_fill = PatternFill(start_color="FF2C3E50", end_color="FF2C3E50", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    data_font = Font(name="Calibri", size=11)
    tag_font = Font(name="Calibri", size=11, bold=True)
    
    thin_border_side = Side(border_style="thin", color="D3D3D3")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    all_processed_rows = []
    total_tags = len(w_groups)

    for idx_tag, (tag, group_data) in enumerate(sorted(w_groups.items(), key=lambda x: sort_tags(x[0]))):
        if progress_callback:
            progress_callback(f"Processing sheet tag '{tag}' ({idx_tag+1}/{total_tags})...", 30 + int(50 * (idx_tag + 1) / total_tags))

        clean_rows = []
        for d in group_data:
            row_dict = d.copy()
            sheet_src = row_dict.pop("Sheet_Source", "Unknown")
            table_id = str(row_dict.get("Table ID", "")).strip()

            genotype, sex = match_table_id(table_id, mapping_dict)
            
            # Insert metadata fields
            row_dict["Timepoint"] = tag
            row_dict["Genotype"] = genotype if genotype else "Unmapped"
            row_dict["Sex"] = sex if sex else "Unmapped"
            row_dict["Group"] = f"{row_dict['Genotype']} {row_dict['Sex']}".strip()

            clean_rows.append(row_dict)
            all_processed_rows.append(row_dict)

        df = pd.DataFrame(clean_rows)

        # Reorder columns logically: Table ID, Genotype, Sex, Group, then metrics
        metadata_cols = ['Table ID', 'Genotype', 'Sex', 'Group']
        metric_cols = [c for c in df.columns if c not in metadata_cols + ['Timepoint']]
        df = df[metadata_cols + metric_cols]

        ws = wb.create_sheet(title=str(tag))

        # Write header & data
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.font = data_font
                cell.border = cell_border

                if r_idx == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    if isinstance(value, float):
                        cell.number_format = '0.000'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

        # Apply visual grouping styling
        if mapping_dict and visual_mode != MODE_COLUMNS_ONLY:
            genotype_col_idx = list(df.columns).index("Genotype") + 1
            sex_col_idx = list(df.columns).index("Sex") + 1
            group_col_idx = list(df.columns).index("Group") + 1

            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                geno_val = row[genotype_col_idx - 1].value
                sex_val = row[sex_col_idx - 1].value

                if geno_val != "Unmapped":
                    fill, tag_font_color = _get_pastel_fill(geno_val, sex_val, genotype_index)

                    if fill:
                        if visual_mode == MODE_PASTEL_ACCENTS:
                            # Apply fill & bold font only to Genotype and Group tag cells
                            row[genotype_col_idx - 1].fill = fill
                            row[genotype_col_idx - 1].font = tag_font_color
                            row[group_col_idx - 1].fill = fill
                            row[group_col_idx - 1].font = tag_font_color
                        elif visual_mode == MODE_SOFT_ROW_FILLS:
                            # Apply soft fill across entire row
                            for c in row:
                                c.fill = fill

        # Auto-adjust column widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Generate Consolidated Summary Per Animal / Timepoint Sheet
    if generate_summary and all_processed_rows:
        if progress_callback:
            progress_callback("Generating consolidated Summary per Animal sheet...", 85)

        df_all = pd.DataFrame(all_processed_rows)

        # Columns structure: Table ID, Timepoint, Genotype, Sex, Group, then all metrics
        lead_cols = ['Table ID', 'Timepoint', 'Genotype', 'Sex', 'Group']
        metric_cols = [c for c in df_all.columns if c not in lead_cols]

        # Group by Table ID and Timepoint to aggregate mean measurements per animal per timepoint
        grouped = df_all.groupby(['Table ID', 'Timepoint'], as_index=False, sort=False)
        
        summary_rows = []
        for (table_id, tp), group_df in grouped:
            geno = group_df['Genotype'].iloc[0]
            sex = group_df['Sex'].iloc[0]
            grp = group_df['Group'].iloc[0]

            row_dict = {
                'Table ID': table_id,
                'Timepoint': tp,
                'Genotype': geno,
                'Sex': sex,
                'Group': grp
            }

            for m in metric_cols:
                vals = group_df[m].dropna().values
                row_dict[m] = float(np.mean(vals)) if len(vals) > 0 else None

            summary_rows.append(row_dict)

        df_summary = pd.DataFrame(summary_rows)
        df_summary = df_summary[lead_cols + metric_cols]

        ws_summary = wb.create_sheet(title="Summary_Per_Animal", index=0)

        # Style header
        sum_header_fill = PatternFill(start_color="FF1A5276", end_color="FF1A5276", fill_type="solid")

        for r_idx, row in enumerate(dataframe_to_rows(df_summary, index=False, header=True), start=1):
            for c_idx, value in enumerate(row, start=1):
                cell = ws_summary.cell(row=r_idx, column=c_idx, value=value)
                cell.font = data_font
                cell.border = cell_border

                if r_idx == 1:
                    cell.fill = sum_header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    if isinstance(value, float):
                        cell.number_format = '0.000'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

        # Apply pastel fills to Genotype/Group tag cells if available
        if mapping_dict and visual_mode != MODE_COLUMNS_ONLY:
            geno_col_idx = lead_cols.index("Genotype") + 1
            sex_col_idx = lead_cols.index("Sex") + 1
            grp_col_idx = lead_cols.index("Group") + 1

            for row in ws_summary.iter_rows(min_row=2, max_row=ws_summary.max_row):
                geno_val = row[geno_col_idx - 1].value
                sex_val = row[sex_col_idx - 1].value

                if geno_val != "Unmapped":
                    fill, tag_font_color = _get_pastel_fill(geno_val, sex_val, genotype_index)
                    if fill:
                        if visual_mode == MODE_PASTEL_ACCENTS:
                            row[geno_col_idx - 1].fill = fill
                            row[geno_col_idx - 1].font = tag_font_color
                            row[grp_col_idx - 1].fill = fill
                            row[grp_col_idx - 1].font = tag_font_color
                        elif visual_mode == MODE_SOFT_ROW_FILLS:
                            for c in row:
                                c.fill = fill

        for col in ws_summary.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_summary.column_dimensions[col_letter].width = max(max_len + 3, 14)

    if progress_callback:
        progress_callback("Saving workbook to disk...", 95)

    wb.save(output_file)

    if progress_callback:
        progress_callback("Done!", 100)

    return output_file
