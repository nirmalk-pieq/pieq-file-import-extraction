#!/usr/bin/env python3
import os
import sys
import argparse
from typing import Union, List
import pandas as pd
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
    # Fallback if no extension
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

def extract_unique(file_path: str, column_name: str, select_cols: List[str] = None,
                   output_path: str = None, sheet_name: Union[str, int, None] = 0):
    """
    Extract unique rows based on a key column and save to file.
    Move rows with empty/NaN key values to a separate file.
    """
    if not os.path.exists(file_path):
        print(f"ERROR: File not found: {file_path}", file=sys.stderr)
        sys.exit(1)

    print(f"Reading {file_path}...")
    df = load_file(file_path, sheet_name=sheet_name)

    # Validate column exists
    if column_name not in df.columns:
        cols_lower = {str(c).strip().lower(): c for c in df.columns}
        matched_col = cols_lower.get(column_name.strip().lower())
        if matched_col:
            print(f"WARNING: Exact column '{column_name}' not found. Using matched column '{matched_col}'.")
            column_name = matched_col
        else:
            print(f"ERROR: Uniqueness column '{column_name}' not found in the file.", file=sys.stderr)
            print(f"Available columns are: {list(df.columns)}", file=sys.stderr)
            sys.exit(1)

    # Validate select columns if provided
    final_cols = list(df.columns)
    if select_cols:
        cleaned_select = []
        missing_cols = []
        for col in select_cols:
            col_strip = col.strip()
            # Try case-insensitive matching
            cols_lower = {str(c).strip().lower(): c for c in df.columns}
            matched = cols_lower.get(col_strip.lower())
            if matched:
                cleaned_select.append(matched)
            else:
                missing_cols.append(col_strip)
        
        if missing_cols:
            print(f"ERROR: Selected output column(s) {missing_cols} not found in the file.", file=sys.stderr)
            print(f"Available columns are: {list(df.columns)}", file=sys.stderr)
            sys.exit(1)
        
        # Ensure the key column is included in the output selection
        if column_name not in cleaned_select:
            print(f"INFO: Auto-adding key column '{column_name}' to output selection.")
            cleaned_select.insert(0, column_name)
            
        final_cols = cleaned_select

    # Determine input filename and extension
    input_dir = os.path.dirname(os.path.abspath(file_path))
    input_base, input_ext = os.path.splitext(os.path.basename(file_path))

    # Resolve output path
    if not output_path:
        output_path = os.path.join(input_dir, f"{input_base}_unique{input_ext}")
    else:
        # If output_path is an existing directory, or ends with a slash, treat it as a directory
        if os.path.isdir(output_path) or output_path.endswith(os.sep) or output_path.endswith('/'):
            output_path = os.path.join(output_path, f"{input_base}_unique{input_ext}")
        else:
            # Check if it has a file extension
            _, out_ext = os.path.splitext(output_path)
            if not out_ext:
                # If no extension, default to the input file's extension
                output_path = output_path + input_ext

    # Determine paths
    output_dir = os.path.dirname(os.path.abspath(output_path))
    try:
        os.makedirs(output_dir, exist_ok=True)
    except Exception as e:
        print(f"ERROR: Could not create output directory '{output_dir}': {e}", file=sys.stderr)
        sys.exit(1)

    # Separate rows with empty / NaN key values
    # Empty string or whitespace-only is also considered empty
    is_empty_mask = df[column_name].isna() | (df[column_name].astype(str).str.strip() == "")
    df_valid = df[~is_empty_mask]
    df_empty = df[is_empty_mask]

    # Process unique valid rows
    total_valid = len(df_valid)
    df_unique = df_valid.drop_duplicates(subset=[column_name], keep='first')
    df_unique_filtered = df_unique[final_cols]
    save_file(df_unique_filtered, output_path)

    # Process empty/NaN key value rows
    empty_output_path = None
    if len(df_empty) > 0:
        base_out, ext_out = os.path.splitext(output_path)
        empty_output_path = f"{base_out}_unassigned{ext_out}"
        df_empty_filtered = df_empty[final_cols]
        save_file(df_empty_filtered, empty_output_path)

    # Print tabular summary
    headers = ["Output Type", "File Name", "Rows", "Full Path"]
    rows = []
    unique_filename = os.path.basename(output_path)
    rows.append(["Unique Entries", unique_filename, str(len(df_unique_filtered)), output_path])
    
    if len(df_empty) > 0:
        empty_filename = os.path.basename(empty_output_path)
        rows.append(["Unassigned Entries", empty_filename, str(len(df_empty)), empty_output_path])
        
    print("\nExtraction Summary:")
    print_ascii_table(headers, rows)

def main():
    parser = argparse.ArgumentParser(
        description="Extract unique rows based on a column and select specific output columns."
    )
    parser.add_argument(
        "-f", "--file",
        required=True,
        help="Path to the input CSV or Excel file."
    )
    parser.add_argument(
        "-c", "--column",
        required=True,
        help="Column name to determine uniqueness by."
    )
    parser.add_argument(
        "-s", "--select",
        default=None,
        help="Comma-separated list of columns to include in the output (e.g. 'Agent Name,Agent ID,Email')."
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Output file path. Defaults to '{input_filename}_unique.{ext}'."
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="Sheet name or index to read (only applicable for Excel files). Defaults to first sheet (0)."
    )

    args = parser.parse_args()

    select_cols = None
    if args.select:
        select_cols = [c.strip() for c in args.select.split(",") if c.strip()]

    sheet = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    extract_unique(
        file_path=args.file,
        column_name=args.column,
        select_cols=select_cols,
        output_path=args.output,
        sheet_name=sheet
    )

if __name__ == "__main__":
    main()
