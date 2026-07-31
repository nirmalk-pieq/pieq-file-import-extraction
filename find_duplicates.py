#!/usr/bin/env python3
"""
Duplicate Row Finder Utility

Finds and marks duplicate rows in an Excel (.xlsx, .xls) or CSV (.csv) file
based on dynamic user-specified column(s).
Adds a 'Duplicates' column ('Yes' if duplicate, 'No' if unique).
Optionally exports duplicate-only rows into a separate file.
"""

import argparse
import os
import sys
import time
import pandas as pd


def clean_key(series: pd.Series) -> pd.Series:
    """Normalizes string values for duplicate comparison (trim, uppercase, collapse spaces)."""
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.strip("'")
        .str.strip('"')
        .str.upper()
        .apply(lambda s: " ".join(s.split()))
    )


def match_column_name(target_col: str, df_columns: list) -> str:
    """Case and space insensitive matching for column names in dataframe."""
    clean_target = " ".join(target_col.strip().lower().split())
    for col in df_columns:
        clean_c = " ".join(str(col).strip().lower().split())
        if clean_c == clean_target:
            return col
    return None


def find_and_mark_duplicates(
    file_path: str,
    target_columns: list,
    output_file_path: str = None,
    separate_duplicates: bool = False,
    sort_columns: list = None,
    sheet_name: str = "0",
) -> str:
    """Finds duplicate rows based on target_columns and marks them in a 'Duplicates' column."""
    start_time = time.time()
    print(f"\n🚀 Starting Duplicate Search in '{os.path.basename(file_path)}'...\n")

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")

    # 1. Load Data File
    print("[1/3] Loading file...", end="", flush=True)
    _, ext = os.path.splitext(file_path.lower())
    try:
        if ext in ('.xlsx', '.xls'):
            sheet_val = int(sheet_name) if sheet_name.isdigit() else sheet_name
            df = pd.read_excel(file_path, sheet_name=sheet_val)
        elif ext == '.csv':
            df = pd.read_csv(file_path)
        else:
            raise ValueError(f"Unsupported file extension '{ext}'. Only Excel (.xlsx, .xls) and CSV (.csv) are supported.")
    except Exception as e:
        print(f" Failed\n❌ Error loading file: {e}", file=sys.stderr)
        sys.exit(1)

    print(f" Done ({len(df):,} rows)")

    # 2. Match Target Columns
    print(f"[2/3] Checking duplicates on column(s): {target_columns}...", end="", flush=True)
    actual_cols = []
    missing_cols = []

    for col in target_columns:
        matched = match_column_name(col, df.columns)
        if matched:
            actual_cols.append(matched)
        else:
            missing_cols.append(col)

    if missing_cols:
        print(f" Failed\n❌ Error: Column(s) {missing_cols} not found in file.", file=sys.stderr)
        print(f"Available columns: {list(df.columns)}", file=sys.stderr)
        sys.exit(1)

    # Handle Sorting if requested
    if sort_columns is not None:
        cols_to_sort = sort_columns if len(sort_columns) > 0 else target_columns
        actual_sort_cols = []
        missing_sort_cols = []
        for col in cols_to_sort:
            matched = match_column_name(col, df.columns)
            if matched:
                actual_sort_cols.append(matched)
            else:
                missing_sort_cols.append(col)

        if missing_sort_cols:
            print(f" Failed\n❌ Error: Sort column(s) {missing_sort_cols} not found in file.", file=sys.stderr)
            sys.exit(1)

        print(f"\nℹ️ Sorting dataset by column(s): {actual_sort_cols}...", end="", flush=True)
        df = df.sort_values(by=actual_sort_cols, ascending=True).reset_index(drop=True)

    # Build normalized keys for duplicate evaluation
    norm_col_names = []
    for c in actual_cols:
        temp_norm_name = f"_norm_{c}"
        df[temp_norm_name] = clean_key(df[c])
        norm_col_names.append(temp_norm_name)

    # Detect duplicates across all matching normalized columns (keep=False marks ALL duplicate rows)
    dup_mask = df.duplicated(subset=norm_col_names, keep=False)
    df["Duplicates"] = dup_mask.map({True: "Yes", False: "No"})

    # Clean up temporary normalized columns
    df = df.drop(columns=norm_col_names)
    print(" Done")

    # 3. Determine Output File Paths
    out_dir = os.path.dirname(output_file_path) if output_file_path else os.path.dirname(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    out_ext = os.path.splitext(output_file_path)[1] if output_file_path else ext

    if not out_ext or out_ext not in ('.xlsx', '.xls', '.csv'):
        out_ext = ext if ext in ('.xlsx', '.xls', '.csv') else '.xlsx'

    if separate_duplicates:
        if not output_file_path:
            clean_output_path = os.path.join(out_dir, f"{base_name}_clean{out_ext}") if out_dir else f"{base_name}_clean{out_ext}"
        else:
            clean_output_path = output_file_path
        dup_file_path = os.path.join(out_dir, f"{base_name}_only_duplicates{out_ext}") if out_dir else f"{base_name}_only_duplicates{out_ext}"
    else:
        if not output_file_path:
            clean_output_path = os.path.join(out_dir, f"{base_name}_duplicates{out_ext}") if out_dir else f"{base_name}_duplicates{out_ext}"
        else:
            clean_output_path = output_file_path
        dup_file_path = None

    out_directory = os.path.dirname(clean_output_path)
    if out_directory:
        os.makedirs(out_directory, exist_ok=True)

    # 4. Save Output File(s)
    print("[3/3] Exporting output file(s)...", end="", flush=True)

    if separate_duplicates:
        clean_df = df[df["Duplicates"] == "No"].drop(columns=["Duplicates"]).copy()
        dup_df = df[df["Duplicates"] == "Yes"].copy()

        if out_ext == '.csv':
            clean_df.to_csv(clean_output_path, index=False)
            dup_df.to_csv(dup_file_path, index=False)
        else:
            clean_df.to_excel(clean_output_path, index=False)
            dup_df.to_excel(dup_file_path, index=False)
        print(" Done")
        print(f"   🟢 Saved clean records (non-duplicates): '{clean_output_path}' ({len(clean_df):,} rows)")
        print(f"   🔴 Saved duplicate records:             '{dup_file_path}' ({len(dup_df):,} rows)")
    else:
        if out_ext == '.csv':
            df.to_csv(clean_output_path, index=False)
        else:
            df.to_excel(clean_output_path, index=False)
        print(" Done")

    # Execution Summary
    elapsed_time = time.time() - start_time
    total_rows = len(df)
    dup_count = int(dup_mask.sum())
    unique_count = total_rows - dup_count
    dup_pct = (dup_count / total_rows * 100.0) if total_rows > 0 else 0.0
    unique_pct = (unique_count / total_rows * 100.0) if total_rows > 0 else 0.0

    print("\n" + "=" * 74)
    print("📊 DUPLICATE SEPARATION SUMMARY")
    print("=" * 74)
    print(f" Total Rows Evaluated:       {total_rows:,}")
    print(f" 🟢 Unique (Non-Duplicate):   {unique_count:,} ({unique_pct:.2f}%)")
    print(f" 🔴 Duplicate Records:        {dup_count:,} ({dup_pct:.2f}%)")
    print(f" Evaluated Column(s):        {actual_cols}")
    if sort_columns is not None:
        print(f" 🔀 Sorted By Column(s):      {actual_sort_cols}")
    if separate_duplicates:
        print(f" 🟢 Clean File (No Dups):     {clean_output_path}")
        print(f" 🔴 Duplicates File:          {dup_file_path}")
    else:
        print(f" 📁 Annotated Output File:    {clean_output_path}")
    print(f" Elapsed Time:               {elapsed_time:.1f}s")
    print("=" * 74 + "\n")

    return clean_output_path


def main():
    parser = argparse.ArgumentParser(
        description="Finds and marks duplicate rows in Excel/CSV files based on dynamic user-selected column(s)."
    )
    parser.add_argument(
        "positional_file",
        nargs="?",
        default=None,
        help="Path to input Excel (.xlsx, .xls) or CSV (.csv) file (alternative to --file).",
    )
    parser.add_argument(
        "--file",
        "-f",
        type=str,
        default=None,
        help="Path to input Excel (.xlsx, .xls) or CSV (.csv) file.",
    )
    parser.add_argument(
        "--columns",
        "-c",
        nargs="+",
        required=True,
        help="One or more column names to group & check duplicates on (e.g. -c 'Policy No' 'Product').",
    )
    parser.add_argument(
        "--output-file",
        "-o",
        type=str,
        default=None,
        help="Path to save annotated output file (default: {file_name}_duplicates.{ext}).",
    )
    parser.add_argument(
        "--separate",
        "-s",
        action="store_true",
        help="Separate duplicate-only rows into an extra file ({file_name}_only_duplicates.{ext}).",
    )
    parser.add_argument(
        "--sort",
        "--sort-by",
        "-sort",
        dest="sort_by",
        nargs="*",
        default=None,
        help="Sort dataset by column(s) before processing. If specified without column names, sorts by evaluation columns (-c).",
    )
    parser.add_argument(
        "--sheet",
        "-sh",
        type=str,
        default="0",
        help="Sheet name or 0-indexed integer for Excel files (default: '0').",
    )

    args = parser.parse_args()

    file_path = args.file or args.positional_file
    if not file_path:
        parser.error("the following arguments are required: --file/-f")

    try:
        find_and_mark_duplicates(
            file_path=file_path,
            target_columns=args.columns,
            output_file_path=args.output_file,
            separate_duplicates=args.separate,
            sort_columns=args.sort_by,
            sheet_name=args.sheet,
        )
    except Exception as e:
        print(f"\n❌ Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
