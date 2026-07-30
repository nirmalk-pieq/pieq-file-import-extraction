#!/usr/bin/env python3
import os
import sys
import argparse
from typing import Union, List, Tuple
from datetime import datetime
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

def calculate_month_range(val, ref_date: datetime) -> Tuple[Union[int, str], Union[int, str]]:
    """
    Calculate From Month and To Month based on month difference relative to ref_date:
    - If 1 to 12 months: From Month = 1, To Month = 12
    - If > 12 months: From Month = 13, To Month = "" (blank)
    - If missing or invalid date: From Month = "", To Month = "" (blank)
    """
    if pd.isna(val):
        return "", ""
    try:
        dt = pd.to_datetime(val, errors='coerce')
        if pd.isna(dt):
            return "", ""
        
        # Calculate month difference
        month_diff = (ref_date.year - dt.year) * 12 + (ref_date.month - dt.month)
        
        if month_diff <= 12:
            return 1, 12
        else:
            return 13, ""
    except Exception:
        return "", ""

def process_effective_date(input_path: str, date_column: str = "Effective Date",
                           current_date_str: str = None, output_dir: str = None,
                           in_place: bool = False, sheet_name: Union[str, int, None] = 0):
    """
    Process input file or directory to calculate From Month and To Month from date_column.
    """
    if not os.path.exists(input_path):
        print(f"ERROR: Input file or directory not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Determine reference date
    if current_date_str:
        try:
            ref_date = pd.to_datetime(current_date_str)
        except Exception as e:
            print(f"ERROR: Invalid --current-date format '{current_date_str}': {e}", file=sys.stderr)
            sys.exit(1)
    else:
        ref_date = datetime.now()

    print(f"Reference Date used for calculation: {ref_date.strftime('%Y-%m-%d')}")

    # Locate files to process
    target_files = []
    if os.path.isdir(input_path):
        for root, _, files in os.walk(input_path):
            for file in files:
                if file.startswith("~$") or file.startswith("."):
                    continue
                _, ext = os.path.splitext(file.lower())
                if ext in ('.xlsx', '.xls', '.csv'):
                    target_files.append(os.path.join(root, file))
    else:
        target_files.append(input_path)

    if not target_files:
        print(f"ERROR: No valid Excel/CSV files found at '{input_path}'.", file=sys.stderr)
        sys.exit(1)

    # Determine destination directory
    if in_place:
        print("Mode: In-place update (overwriting original target files)")
    else:
        if not output_dir:
            if os.path.isdir(input_path):
                output_dir = input_path.rstrip(os.sep) + "_processed"
            else:
                base, _ = os.path.splitext(input_path)
                output_dir = os.path.dirname(os.path.abspath(input_path))
        os.makedirs(output_dir, exist_ok=True)
        print(f"Mode: Saving processed files to directory: {output_dir}")

    summary_rows = []
    for t_file in target_files:
        t_name = os.path.basename(t_file)
        try:
            df = load_file(t_file, sheet_name=sheet_name)
        except Exception as e:
            print(f"WARNING: Skipping '{t_name}' (failed to load: {e})", file=sys.stderr)
            continue

        # Case & whitespace insensitive column matching
        cols_lower = {str(c).strip().lower(): c for c in df.columns}
        matched_date_col = cols_lower.get(date_column.strip().lower())

        if not matched_date_col:
            print(f"WARNING: Date column '{date_column}' not found in '{t_name}'. Skipping.", file=sys.stderr)
            continue

        # Calculate From Month and To Month for each row
        from_months = []
        to_months = []
        for val in df[matched_date_col]:
            fm, tm = calculate_month_range(val, ref_date)
            from_months.append(fm)
            to_months.append(tm)

        df["From Month"] = from_months
        df["To Month"] = to_months

        # Determine output file path
        if in_place:
            out_file_path = t_file
        else:
            if os.path.isdir(input_path):
                rel_path = os.path.relpath(t_file, input_path)
                out_file_path = os.path.join(output_dir, rel_path)
                os.makedirs(os.path.dirname(out_file_path), exist_ok=True)
            else:
                base_t, ext_t = os.path.splitext(t_name)
                out_file_path = os.path.join(output_dir, f"{base_t}_processed{ext_t}")

        save_file(df, out_file_path)
        summary_rows.append([t_name, str(len(df)), matched_date_col, out_file_path])

    print("\nProcessing Summary:")
    print_ascii_table(["File Name", "Total Rows", "Matched Date Column", "Output Path"], summary_rows)

def main():
    parser = argparse.ArgumentParser(
        description="Calculate month differences from Effective Date and populate From Month and To Month columns."
    )
    parser.add_argument(
        "-f", "--file",
        required=True,
        help="Path to an input Excel/CSV file OR directory containing target files."
    )
    parser.add_argument(
        "-c", "--date-column",
        default="Effective Date",
        help="Column name for the effective date. Defaults to 'Effective Date'."
    )
    parser.add_argument(
        "--current-date",
        default=None,
        help="Optional reference date in YYYY-MM-DD format. Defaults to current system date."
    )
    parser.add_argument(
        "-o", "--output-dir",
        default=None,
        help="Output directory to save processed files. Defaults to '{input}_processed'."
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Overwrite original files directly instead of saving to a new directory."
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="Sheet name or index to read for Excel files. Defaults to first sheet (0)."
    )

    args = parser.parse_args()

    sheet = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    process_effective_date(
        input_path=args.file,
        date_column=args.date_column,
        current_date_str=args.current_date,
        output_dir=args.output_dir,
        in_place=args.in_place,
        sheet_name=sheet
    )

if __name__ == "__main__":
    main()
