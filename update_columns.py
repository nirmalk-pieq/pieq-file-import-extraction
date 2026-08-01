#!/usr/bin/env python3
import os
import sys
import argparse
from typing import Union, List, Dict
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def load_file(file_path: str, sheet_name: Union[str, int, None] = 0) -> pd.DataFrame:
    """
    Load the input Excel or CSV file into a pandas DataFrame.
    """
    _, ext = os.path.splitext(file_path.lower())
    if ext in ('.xlsx', '.xls'):
        try:
            return pd.read_excel(file_path, sheet_name=sheet_name)
        except Exception as e:
            print(f"ERROR: Failed to read Excel file '{file_path}': {e}", file=sys.stderr)
            sys.exit(1)
    elif ext == '.csv':
        try:
            return pd.read_csv(file_path)
        except Exception as e:
            print(f"ERROR: Failed to read CSV file '{file_path}': {e}", file=sys.stderr)
            sys.exit(1)
    else:
        print(f"ERROR: Unsupported file extension '{ext}'. Only Excel (.xlsx, .xls) and CSV (.csv) are supported.", file=sys.stderr)
        sys.exit(1)

def save_excel_with_format(df: pd.DataFrame, file_path: str):
    """
    Save DataFrame to Excel and apply professional formatting using xlsxwriter.
    Guarantees sharedStrings.xml creation for compatibility with parsers like pylightxl.
    """
    try:
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Ensure gridlines are visible
            worksheet.hide_gridlines(2)
            
            # Header format (bold, grey background, thin bottom border)
            header_format = workbook.add_format({
                'bold': True,
                'font_name': 'Calibri',
                'font_size': 11,
                'font_color': '#333333',
                'bg_color': '#F2F2F2',
                'align': 'left',
                'valign': 'vcenter',
                'bottom': 1,
                'bottom_color': '#CCCCCC'
            })
            
            # Data cells formatting (explicitly plain Calibri 10)
            data_format = workbook.add_format({
                'bold': False,
                'font_name': 'Calibri',
                'font_size': 10,
                'font_color': '#000000',
                'align': 'left',
                'valign': 'vcenter'
            })
            
            # Re-write headers explicitly to apply format
            for col_idx, col_name in enumerate(df.columns):
                worksheet.write(0, col_idx, col_name, header_format)
                
            # Set column widths and data formats
            for col_idx, col_name in enumerate(df.columns):
                max_val_len = 0
                if len(df) > 0:
                    max_val_len = int(df[col_name].apply(lambda x: len(str(x)) if pd.notna(x) else 0).max())
                max_len = max(max_val_len, len(str(col_name)))
                col_width = int(min(max(max_len + 3, 10), 50))
                worksheet.set_column(col_idx, col_idx, col_width, data_format)
    except Exception as e:
        print(f"WARNING: xlsxwriter formatting failed, falling back: {e}", file=sys.stderr)
        df.to_excel(file_path, index=False)

def print_ascii_table(headers: List[str], rows: List[List[str]]):
    """
    Prints a list of rows as a beautiful, dense ASCII table.
    """
    if not rows:
        return
    
    # Calculate column widths
    widths = [len(h) for h in headers]
    for r in rows:
        for i, val in enumerate(r):
            widths[i] = max(widths[i], len(str(val)))
            
    # Border generator
    border = "+" + "+".join(["-" * (w + 2) for w in widths]) + "+"
    
    print(border)
    # Header
    header_line = "|" + "|".join([f" {headers[i].ljust(widths[i])} " for i in range(len(headers))]) + "|"
    print(header_line)
    print(border)
    
    # Rows
    for r in rows:
        row_line = "|" + "|".join([f" {str(r[i]).ljust(widths[i])} " for i in range(len(r))]) + "|"
        print(row_line)
        
    print(border)

def save_file(df: pd.DataFrame, output_path: str):
    """
    Save DataFrame to CSV or Excel based on the file extension and format if Excel.
    """
    _, ext = os.path.splitext(output_path.lower())
    if not ext:
        ext = '.xlsx'
        output_path = output_path + ext

    try:
        if ext in ('.xlsx', '.xls'):
            save_excel_with_format(df, output_path)
        elif ext == '.csv':
            df.to_csv(output_path, index=False)
        else:
            save_excel_with_format(df, output_path)
    except Exception as e:
        print(f"ERROR: Could not save output to '{output_path}': {e}", file=sys.stderr)
        sys.exit(1)

def is_value_missing(val) -> bool:
    """Checks if a cell value is missing, empty, or an unmatched placeholder (including '\\N')."""
    if pd.isna(val):
        return True
    s = str(val).strip()
    if not s or s.upper() in ("", "NAN", "NONE", "NULL", "[UNMATCHED]", "UNMATCHED", r"\N", r"\\N"):
        return True
    return False

def build_norm_key(df: pd.DataFrame, key_cols: List[str]) -> pd.Series:
    """
    Build a composite normalized key string joined by '||' from multiple columns.
    Normalizes case, outer spaces, and collapses multiple internal spaces.
    """
    def row_to_key(row):
        parts = [" ".join(str(val).strip().lower().split()) for val in row]
        return "||".join(parts)
    return df[key_cols].apply(row_to_key, axis=1)

def update_columns(source_path: str, key_column: Union[str, List[str]], update_cols: List[str],
                   target_path: str, target_key_column: Union[str, List[str]] = None,
                   output_dir: str = None, in_place: bool = False,
                   sheet_name: Union[str, int, None] = 0):
    """
    Match target file(s) with reference source_path on single or composite key_column(s) and update update_cols.
    """
    if not os.path.exists(source_path):
        print(f"ERROR: Source file not found: {source_path}", file=sys.stderr)
        sys.exit(1)

    if not os.path.exists(target_path):
        print(f"ERROR: Target file/folder not found: {target_path}", file=sys.stderr)
        sys.exit(1)

    # Normalize key_column argument to list of strings
    if isinstance(key_column, str):
        raw_key_list = [k.strip() for k in key_column.split(",") if k.strip()]
    else:
        raw_key_list = [str(k).strip() for k in key_column if str(k).strip()]

    # Normalize target_key_column argument to list of strings
    if target_key_column:
        if isinstance(target_key_column, str):
            raw_target_key_list = [k.strip() for k in target_key_column.split(",") if k.strip()]
        else:
            raw_target_key_list = [str(k).strip() for k in target_key_column if str(k).strip()]
    else:
        raw_target_key_list = raw_key_list

    if len(raw_target_key_list) != len(raw_key_list):
        print(f"ERROR: Number of target key columns ({len(raw_target_key_list)}) does not match source key columns ({len(raw_key_list)}).", file=sys.stderr)
        sys.exit(1)

    print(f"Reading reference file: {source_path}...")
    df_source = load_file(source_path, sheet_name=sheet_name)

    # Validate key_column(s) in source
    cols_lower_source = {str(c).strip().lower(): c for c in df_source.columns}
    clean_source_keys = []
    missing_source_keys = []
    for k in raw_key_list:
        matched = cols_lower_source.get(k.lower())
        if matched:
            clean_source_keys.append(matched)
        else:
            missing_source_keys.append(k)

    if missing_source_keys:
        print(f"ERROR: Key column(s) {missing_source_keys} not found in source file.", file=sys.stderr)
        print(f"Available columns in source: {list(df_source.columns)}", file=sys.stderr)
        sys.exit(1)

    # Validate update_cols in source
    clean_update_cols = []
    missing_cols = []
    for col in update_cols:
        col_strip = col.strip()
        matched = cols_lower_source.get(col_strip.lower())
        if matched:
            clean_update_cols.append(matched)
        else:
            missing_cols.append(col_strip)

    if missing_cols:
        print(f"ERROR: Update column(s) {missing_cols} not found in source file.", file=sys.stderr)
        print(f"Available columns in source: {list(df_source.columns)}", file=sys.stderr)
        sys.exit(1)

    # Clean source data: drop NaNs in key columns and keep first occurrence
    df_source_clean = df_source.dropna(subset=clean_source_keys).copy()
    df_source_clean['_norm_key'] = build_norm_key(df_source_clean, clean_source_keys)
    df_source_clean = df_source_clean.drop_duplicates(subset=['_norm_key'], keep='first')

    # Build mapping dictionary for each update column
    mappings: Dict[str, dict] = {}
    for ucol in clean_update_cols:
        mappings[ucol] = dict(zip(df_source_clean['_norm_key'], df_source_clean[ucol]))

    # Locate target files
    target_files = []
    if os.path.isdir(target_path):
        for root, _, files in os.walk(target_path):
            for file in files:
                if file.startswith("~$") or file.startswith("."):
                    continue
                _, ext = os.path.splitext(file.lower())
                if ext in ('.xlsx', '.xls', '.csv'):
                    target_files.append(os.path.join(root, file))
    else:
        target_files.append(target_path)

    if not target_files:
        print(f"ERROR: No valid Excel/CSV target files found in '{target_path}'.", file=sys.stderr)
        sys.exit(1)

    # Determine destination directory/files
    if in_place:
        print("Mode: In-place update (overwriting original target files)")
    else:
        if not output_dir:
            if os.path.isdir(target_path):
                output_dir = target_path.rstrip(os.sep) + "_updated"
            else:
                base, _ = os.path.splitext(target_path)
                output_dir = os.path.dirname(os.path.abspath(target_path))
        os.makedirs(output_dir, exist_ok=True)
        print(f"Mode: Saving updated files to directory: {output_dir}")

    summary_rows = []
    for t_file in target_files:
        t_name = os.path.basename(t_file)
        try:
            df_target = load_file(t_file, sheet_name=sheet_name)
        except Exception as e:
            print(f"WARNING: Skipping '{t_name}' (failed to load: {e})", file=sys.stderr)
            continue

        # Check target key column(s)
        t_cols_lower = {str(c).strip().lower(): c for c in df_target.columns}
        clean_target_keys = []
        missing_t_keys = []
        for tk in raw_target_key_list:
            t_matched = t_cols_lower.get(tk.lower())
            if t_matched:
                clean_target_keys.append(t_matched)
            else:
                missing_t_keys.append(tk)

        if missing_t_keys:
            print(f"WARNING: Key column(s) {missing_t_keys} not found in '{t_name}'. Skipping.", file=sys.stderr)
            continue

        # Create normalized target composite key for matching
        t_norm_key = build_norm_key(df_target, clean_target_keys)

        # Count row matches
        matched_rows_count = t_norm_key.isin(df_source_clean['_norm_key']).sum()

        # Update or insert columns
        for ucol in clean_update_cols:
            mapped_series = t_norm_key.map(mappings[ucol])
            
            # Check if ucol already exists in target (case & whitespace insensitive)
            existing_t_col = t_cols_lower.get(str(ucol).strip().lower())
            if existing_t_col:
                m_ser = mapped_series.astype(object)
                t_ser = df_target[existing_t_col].astype(object)
                
                m_missing = m_ser.apply(is_value_missing)
                
                # Where source is missing (\N, NaN, empty), preserve target value; otherwise use source value
                res = pd.Series(np.where(m_missing, t_ser, m_ser), index=df_target.index)
                df_target[existing_t_col] = res
            else:
                # Insert as new column
                df_target[ucol] = mapped_series

        # Determine output file path
        if in_place:
            out_file_path = t_file
        else:
            if os.path.isdir(target_path):
                rel_path = os.path.relpath(t_file, target_path)
                out_file_path = os.path.join(output_dir, rel_path)
                os.makedirs(os.path.dirname(out_file_path), exist_ok=True)
            else:
                base_t, ext_t = os.path.splitext(t_name)
                out_file_path = os.path.join(output_dir, f"{base_t}_updated{ext_t}")

        save_file(df_target, out_file_path)
        summary_rows.append([t_name, str(len(df_target)), f"{matched_rows_count}/{len(df_target)}", out_file_path])

    print("\nUpdate Summary:")
    print_ascii_table(["Target File", "Total Rows", "Matched Keys", "Output Path"], summary_rows)

def main():
    parser = argparse.ArgumentParser(
        description="Update or add column values in target file(s) by matching single or composite key columns against a reference file."
    )
    parser.add_argument(
        "-s", "--source",
        required=True,
        help="Path to reference lookup Excel/CSV file (e.g. Agent_data.xlsx)."
    )
    parser.add_argument(
        "-k", "--key",
        required=True,
        help="Comma-separated key column name(s) to match on in reference file (e.g. 'Policy No, Agent Name')."
    )
    parser.add_argument(
        "-u", "--update-cols",
        required=True,
        help="Comma-separated list of column(s) from reference file to update/add in target files (e.g. 'Carrier, Agent ID')."
    )
    parser.add_argument(
        "-t", "--target",
        required=True,
        help="Path to a target file OR folder containing target Excel/CSV files."
    )
    parser.add_argument(
        "--target-key",
        default=None,
        help="Comma-separated key column name(s) in target file(s) if different from --key. Defaults to --key."
    )
    parser.add_argument(
        "-o", "--output-dir",
        default=None,
        help="Output directory to save updated file(s). Defaults to '{target}_updated'."
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Overwrite original target file(s) directly instead of saving to a new directory."
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="Sheet name or index to read for Excel files. Defaults to first sheet (0)."
    )

    args = parser.parse_args()

    update_list = [c.strip() for c in args.update_cols.split(",") if c.strip()]

    sheet = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    update_columns(
        source_path=args.source,
        key_column=args.key,
        update_cols=update_list,
        target_path=args.target,
        target_key_column=args.target_key,
        output_dir=args.output_dir,
        in_place=args.in_place,
        sheet_name=sheet
    )

if __name__ == "__main__":
    main()

