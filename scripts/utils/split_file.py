#!/usr/bin/env python3
import os
import sys
import re
import argparse
from typing import Union, List
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def sanitize_filename(name: str) -> str:
    """
    Sanitize the group name to ensure it is a valid filename.
    Replaces common invalid characters with underscores.
    """
    # Replace characters that are invalid in Windows/Mac/Linux filenames
    sanitized = re.sub(r'[\\/*?:"<>|]', '_', name)
    # Strip leading/trailing whitespaces or dots
    sanitized = sanitized.strip().strip('.')
    # Fallback if filename becomes empty
    if not sanitized:
        sanitized = "unassigned_group"
    return sanitized

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

def split_file(file_path: str, column_name: str, output_dir: str = None, 
               out_format: str = 'match', sheet_name: Union[str, int, None] = 0):
    """
    Split the input file by grouping on the specified column_name.
    """
    if os.path.isdir(file_path):
        files = [
            os.path.join(file_path, f) for f in sorted(os.listdir(file_path))
            if f.lower().endswith(('.csv', '.xlsx', '.xls')) and not f.startswith('~$')
        ]
        if not files:
            print(f"ERROR: No CSV or Excel files found in directory '{file_path}'.", file=sys.stderr)
            sys.exit(1)

        print(f"Directory detected. Combining {len(files)} file(s) from '{file_path}'...")
        dfs = [load_file(f, sheet_name=sheet_name) for f in files]
        df = pd.concat(dfs, ignore_index=True, sort=False)
        _, input_ext = os.path.splitext(files[0].lower())
    else:
        print(f"Reading {file_path}...")
        df = load_file(file_path, sheet_name=sheet_name)
        _, input_ext = os.path.splitext(file_path.lower())

    # Validate column exists
    if column_name not in df.columns:
        cols_lower = {str(c).strip().lower(): c for c in df.columns}
        matched_col = cols_lower.get(column_name.strip().lower())
        if matched_col:
            print(f"WARNING: Exact column '{column_name}' not found. Using matched column '{matched_col}'.")
            column_name = matched_col
        else:
            print(f"ERROR: Column '{column_name}' not found in dataset.", file=sys.stderr)
            print(f"Available columns are: {list(df.columns)}", file=sys.stderr)
            sys.exit(1)

    total_rows = len(df)
    print(f"Loaded {total_rows:,} rows. Grouping by column '{column_name}'...")

    # Set default output directory if not provided
    if not output_dir:
        input_dir = os.path.dirname(os.path.abspath(file_path))
        output_dir = os.path.join(input_dir, "split_output")
    
    try:
        os.makedirs(output_dir, exist_ok=True)
    except Exception as e:
        print(f"ERROR: Could not create output directory '{output_dir}': {e}", file=sys.stderr)
        sys.exit(1)
    print(f"Outputs will be saved in: {output_dir}")

    # Group by the specified column
    # Handle NaN values explicitly by replacing them with a string identifier
    # Note: df[column_name] might contain NaNs. groupby ignores NaNs by default or puts them under NaN.
    # To ensure we capture them and save to a file, we fillna first or group with dropna=False (pandas >= 1.1.0)
    grouped = df.groupby(column_name, dropna=False)

    files_created = 0
    table_rows = []
    for group_val, group_df in grouped:
        # Determine output format and extension
        if out_format == 'match':
            target_ext = input_ext if input_ext in ('.xlsx', '.xls', '.csv') else '.xlsx'
        elif out_format == 'csv':
            target_ext = '.csv'
        else:
            target_ext = '.xlsx'

        # Check if the group value is NaN/Null
        if pd.isna(group_val):
            group_str = f"unassigned_{column_name}"
        else:
            group_str = str(group_val).strip()
            if not group_str:
                group_str = f"empty_{column_name}"

        sanitized_name = sanitize_filename(group_str)
        filename = f"{sanitized_name}{target_ext}"
        output_path = os.path.join(output_dir, filename)

        # Save files
        save_file(group_df, output_path)
        files_created += 1
        table_rows.append([filename, str(len(group_df)), output_path])

    print("\nSplit Summary:")
    print_ascii_table(["File Name", "Rows", "Full Path"], table_rows)

def main():
    parser = argparse.ArgumentParser(
        description="Split a CSV or Excel file into multiple files based on unique values in a specified column."
    )
    parser.add_argument(
        "-f", "--file",
        required=True,
        help="Path to the input CSV or Excel file."
    )
    parser.add_argument(
        "-c", "--column",
        required=True,
        help="Column name to group and split by."
    )
    parser.add_argument(
        "-o", "--output-dir",
        default=None,
        help="Directory to save the split files. Defaults to a 'split_output' folder in the input file's directory."
    )
    parser.add_argument(
        "--format",
        choices=["match", "excel", "csv"],
        default="match",
        help="Output file format (excel: .xlsx, csv: .csv, match: match the input file type). Default: match."
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="Sheet name or index to read (only applicable for Excel files). Defaults to the first sheet (0)."
    )

    args = parser.parse_args()

    # Parse sheet argument: convert digits to int if appropriate
    sheet = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    split_file(
        file_path=args.file,
        column_name=args.column,
        output_dir=args.output_dir,
        out_format=args.format,
        sheet_name=sheet
    )

if __name__ == "__main__":
    main()
