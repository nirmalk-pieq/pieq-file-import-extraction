#!/usr/bin/env python3
"""
Join / Concatenate Policy Files

Utility script to combine multiple Excel or CSV files (e.g., clean output + error output)
into a single consolidated file.
"""

import os
import sys
import argparse
from typing import List
import pandas as pd

def load_file(file_path: str) -> pd.DataFrame:
    """Load an Excel or CSV file into a pandas DataFrame."""
    if not os.path.exists(file_path):
        print(f"ERROR: File not found: '{file_path}'", file=sys.stderr)
        sys.exit(1)
        
    _, ext = os.path.splitext(file_path.lower())
    try:
        if ext in ('.xlsx', '.xls'):
            return pd.read_excel(file_path)
        elif ext == '.csv':
            return pd.read_csv(file_path)
        else:
            print(f"ERROR: Unsupported file format '{ext}' for file '{file_path}'.", file=sys.stderr)
            sys.exit(1)
    except Exception as e:
        print(f"ERROR: Failed to read '{file_path}': {e}", file=sys.stderr)
        sys.exit(1)

def save_excel_with_format(df: pd.DataFrame, output_path: str):
    """Save DataFrame to Excel with header formatting using xlsxwriter if available."""
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            worksheet.hide_gridlines(2)
            
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
            
            data_format = workbook.add_format({
                'bold': False,
                'font_name': 'Calibri',
                'font_size': 10,
                'font_color': '#000000',
                'align': 'left',
                'valign': 'vcenter'
            })
            
            for col_idx, col_name in enumerate(df.columns):
                worksheet.write(0, col_idx, col_name, header_format)
                
            for col_idx, col_name in enumerate(df.columns):
                max_val_len = 0
                if len(df) > 0:
                    max_val_len = int(df[col_name].apply(lambda x: len(str(x)) if pd.notna(x) else 0).max())
                max_len = max(max_val_len, len(str(col_name)))
                col_width = int(min(max(max_len + 3, 10), 50))
                worksheet.set_column(col_idx, col_idx, col_width, data_format)
    except Exception as e:
        df.to_excel(output_path, index=False)

def join_files(input_paths: List[str], output_path: str):
    """Joins/concatenates multiple Excel/CSV files into a single consolidated file."""
    if not input_paths:
        print("ERROR: No input files provided.", file=sys.stderr)
        sys.exit(1)

    print(f"\n🚀 Joining {len(input_paths)} file(s)...")
    
    dfs = []
    file_summaries = []

    for path in input_paths:
        filename = os.path.basename(path)
        print(f"  • Reading '{filename}'...", end="", flush=True)
        df = load_file(path)
        dfs.append(df)
        file_summaries.append((filename, len(df)))
        print(f" Done ({len(df):,} rows)")

    combined_df = pd.concat(dfs, ignore_index=True, sort=False)
    
    # Save output
    print(f"\n[Exporting combined file to '{output_path}'...]", end="", flush=True)
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    _, ext = os.path.splitext(output_path.lower())
    if ext == '.csv':
        combined_df.to_csv(output_path, index=False)
    else:
        if not (output_path.endswith('.xlsx') or output_path.endswith('.xls')):
            output_path += '.xlsx'
        save_excel_with_format(combined_df, output_path)
    print(" Done")

    # Display execution summary
    print("\n" + "=" * 65)
    print("📊 FILE JOIN SUMMARY")
    print("=" * 65)
    for fname, count in file_summaries:
        print(f" File: {fname:<40} | {count:<12,} rows")
    print("-" * 65)
    print(f" Total Combined Records:                        | {len(combined_df):<12,} rows")
    print("=" * 65)
    print(f"\n✅ Combined file successfully saved to: {output_path}\n")

def main():
    parser = argparse.ArgumentParser(
        description="Join / concatenate multiple Excel or CSV files into a single file."
    )
    parser.add_argument(
        "files",
        nargs="*",
        help="Input Excel/CSV file paths to join (e.g. output.xlsx output_Errors.xlsx)."
    )
    parser.add_argument(
        "-i", "--inputs",
        nargs="+",
        dest="input_files",
        help="Input file paths (alternative to positional arguments)."
    )
    parser.add_argument(
        "-o", "--output",
        required=True,
        help="Path where the output combined Excel/CSV file will be saved."
    )

    args = parser.parse_args()

    input_list = args.files if args.files else []
    if args.input_files:
        input_list.extend(args.input_files)

    # Remove duplicates preserving order
    seen = set()
    unique_inputs = []
    for f in input_list:
        if f not in seen:
            seen.add(f)
            unique_inputs.append(f)

    if not unique_inputs:
        parser.error("At least one input file must be specified.")

    join_files(unique_inputs, args.output)

if __name__ == "__main__":
    main()
