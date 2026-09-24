#!/usr/bin/env python3
"""
Extract Unique Commission Import Script

Looks up Policy Numbers from the Policy Import output (file or directory)
against Commission Structure files (file or directory), extracts matching
unique commission records, and saves them into an output directory.
"""

import os
import sys
import argparse
import time
from typing import List, Set, Tuple, Union
import pandas as pd


def clean_key(series: pd.Series) -> pd.Series:
    """Normalizes lookup keys (strips quotes/whitespace, uppercase, collapses spaces)."""
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.strip("'")
        .str.strip('"')
        .str.upper()
        .apply(lambda s: " ".join(s.split()))
    )


def load_file(file_path: str) -> pd.DataFrame:
    """Loads an Excel or CSV file into a pandas DataFrame."""
    if not os.path.exists(file_path):
        print(f"❌ Error: File not found: '{file_path}'", file=sys.stderr)
        sys.exit(1)

    _, ext = os.path.splitext(file_path.lower())
    try:
        if ext in ('.xlsx', '.xls'):
            return pd.read_excel(file_path)
        elif ext == '.csv':
            return pd.read_csv(file_path)
        else:
            print(f"❌ Error: Unsupported file format '{ext}' for file '{file_path}'.", file=sys.stderr)
            sys.exit(1)
    except Exception as e:
        print(f"❌ Error: Failed to read '{file_path}': {e}", file=sys.stderr)
        sys.exit(1)


def find_policy_col(df: pd.DataFrame, preferred_name: str = None) -> str:
    """Auto-detects or validates Policy Number column name in DataFrame."""
    if preferred_name:
        for c in df.columns:
            if str(c).strip().lower() == str(preferred_name).strip().lower():
                return c
        print(f"⚠️ Warning: Specified policy column '{preferred_name}' not found. Attempting auto-detection...")

    candidates = [
        "policy no", "policy number", "policyno", "policynumber", "policy_no", "policy_number", "pol_no"
    ]
    for c in df.columns:
        c_clean = str(c).strip().lower()
        if c_clean in candidates:
            return c
        
    for c in df.columns:
        c_clean = str(c).strip().lower()
        if "policy" in c_clean and ("no" in c_clean or "num" in c_clean or "id" in c_clean):
            return c

    return None


def collect_input_files(path: str) -> List[str]:
    """Collects single file path or list of Excel/CSV files in a directory."""
    if not os.path.exists(path):
        print(f"❌ Error: Path not found: '{path}'", file=sys.stderr)
        sys.exit(1)

    if os.path.isfile(path):
        return [path]

    files = [
        os.path.join(path, f) for f in sorted(os.listdir(path))
        if f.lower().endswith(('.csv', '.xlsx', '.xls')) and not f.startswith('~$')
    ]
    if not files:
        print(f"❌ Error: No CSV or Excel files found in directory '{path}'.", file=sys.stderr)
        sys.exit(1)
    return files


def extract_policy_keys_from_input(policy_input_path: str, policy_col_override: str = None) -> Set[str]:
    """Loads policy numbers from Policy Import file/folder into a set of normalized keys."""
    policy_files = collect_input_files(policy_input_path)
    valid_keys = set()
    total_rows = 0

    print(f"[1/3] Extracting policy numbers from Policy Import dataset ({len(policy_files)} file(s))...")
    for filepath in policy_files:
        filename = os.path.basename(filepath)
        df = load_file(filepath)
        total_rows += len(df)

        pol_col = find_policy_col(df, policy_col_override)
        if not pol_col:
            print(f"  ⚠️ Warning: Could not find Policy Number column in '{filename}'. Skipping file.")
            continue

        clean_keys = clean_key(df[pol_col])
        non_empty = clean_keys[clean_keys != ""]
        valid_keys.update(non_empty)
        print(f"  • '{filename}': {len(df):,} rows -> {len(set(non_empty)):,} unique policy keys")

    print(f"  ✅ Extracted total {len(valid_keys):,} unique Policy Number(s) across {total_rows:,} evaluated policy row(s).\n")
    return valid_keys


def save_output_file(df: pd.DataFrame, output_path: str):
    """Saves DataFrame to CSV or Excel formatted file."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    _, ext = os.path.splitext(output_path.lower())
    if ext in ('.xlsx', '.xls'):
        df.to_excel(output_path, index=False)
    else:
        df.to_csv(output_path, index=False)


def process_commission_files(
    comm_input_path: str,
    valid_policy_keys: Set[str],
    output_dir: str,
    comm_col_override: str = None,
    dedupe: bool = True,
    separate_unmatched: bool = True,
) -> Tuple[List[pd.DataFrame], List[pd.DataFrame]]:
    """Filters commission structure file(s) against valid_policy_keys, deduplicates, and saves output."""
    comm_files = collect_input_files(comm_input_path)
    is_dir_input = os.path.isdir(comm_input_path)

    print(f"[2/3] Processing Commission Structure dataset ({len(comm_files)} file(s))...")

    matched_dfs = []
    unmatched_dfs = []

    unmatched_out_dir = os.path.join(output_dir, "unmatched_commission") if separate_unmatched else None

    for idx, filepath in enumerate(comm_files, 1):
        filename = os.path.basename(filepath)
        print(f"  ({idx}/{len(comm_files)}) Filtering '{filename}'...", end="", flush=True)

        df = load_file(filepath)
        if len(df) == 0:
            print(" Skipping (0 rows)")
            continue

        pol_col = find_policy_col(df, comm_col_override)
        if not pol_col:
            print(f" FAILED (Policy column not found)")
            continue

        df["_pol_key"] = clean_key(df[pol_col])
        matched_mask = df["_pol_key"].isin(valid_policy_keys)

        matched_df = df[matched_mask].drop(columns=["_pol_key"]).copy()
        unmatched_df = df[~matched_mask].drop(columns=["_pol_key"]).copy()

        raw_matched_count = len(matched_df)

        if dedupe and raw_matched_count > 0:
            key_cols = [c for c in ["Policy No", "Product", "Sequence", "Agent Name", "From Month", "Amount"] if c in matched_df.columns]
            if key_cols:
                matched_df = matched_df.drop_duplicates(subset=key_cols, keep="first")
            else:
                matched_df = matched_df.drop_duplicates(keep="first")

        unique_matched_count = len(matched_df)
        matched_dfs.append(matched_df)
        if len(unmatched_df) > 0:
            unmatched_dfs.append(unmatched_df)

        # Save individual file if input was directory or single file specified
        raw_base, ext = os.path.splitext(filename)
        out_ext = ".csv" if ext.lower() == ".csv" else ".xlsx"
        
        if is_dir_input:
            out_file_name = f"{raw_base}_unique{out_ext}"
            out_filepath = os.path.join(output_dir, out_file_name)
            save_output_file(matched_df, out_filepath)

        if separate_unmatched and len(unmatched_df) > 0 and unmatched_out_dir:
            unmatched_filepath = os.path.join(unmatched_out_dir, f"{raw_base}_unmatched{out_ext}")
            save_output_file(unmatched_df, unmatched_filepath)

        print(f" Done ({raw_matched_count:,} matched -> {unique_matched_count:,} unique rows | {len(unmatched_df):,} unmatched)")

    return matched_dfs, unmatched_dfs


def main():
    parser = argparse.ArgumentParser(
        description="Extract unique commission import records matching Policy Import policy numbers."
    )
    parser.add_argument(
        "--policy-input",
        "--policy-file",
        "--policy-dir",
        "--policy-path",
        dest="policy_input",
        type=str,
        required=True,
        help="Path to Policy Import output file (.xlsx/.csv) OR folder containing policy files.",
    )
    parser.add_argument(
        "--commission-input",
        "--commission-file",
        "--commission-dir",
        "--commission-path",
        dest="commission_input",
        type=str,
        required=True,
        help="Path to Commission Structure file (.xlsx/.csv) OR folder containing commission files.",
    )
    parser.add_argument(
        "--output-dir",
        "--output-folder",
        dest="output_dir",
        type=str,
        default="pipeline_output/unique_commission_import",
        help="Target directory to save unique commission import file(s).",
    )
    parser.add_argument(
        "--policy-key-col",
        type=str,
        default=None,
        help="Optional Policy Number column name in policy import files (default: auto-detect).",
    )
    parser.add_argument(
        "--comm-key-col",
        type=str,
        default=None,
        help="Optional Policy Number column name in commission structure files (default: auto-detect).",
    )
    parser.add_argument(
        "--no-dedupe",
        dest="dedupe",
        action="store_false",
        help="Disable deduplication of matching commission rows.",
    )
    parser.add_argument(
        "--no-separate-unmatched",
        dest="separate_unmatched",
        action="store_false",
        help="Disable saving unmatched commission rows to unmatched_commission/ subfolder.",
    )

    args = parser.parse_args()
    start_time = time.time()

    print("🚀 Starting Unique Commission Import Extraction...\n")

    # 1. Load valid policy keys
    valid_keys = extract_policy_keys_from_input(args.policy_input, args.policy_key_col)
    if not valid_keys:
        print("❌ Error: No valid policy keys found in Policy Import dataset.", file=sys.stderr)
        sys.exit(1)

    # 2. Process commission files
    matched_dfs, unmatched_dfs = process_commission_files(
        comm_input_path=args.commission_input,
        valid_policy_keys=valid_keys,
        output_dir=args.output_dir,
        comm_col_override=args.comm_key_col,
        dedupe=args.dedupe,
        separate_unmatched=args.separate_unmatched,
    )

    # 3. Export combined output file
    print("\n[3/3] Consolidating output datasets...", end="", flush=True)
    if matched_dfs:
        combined_matched = pd.concat(matched_dfs, ignore_index=True)
    else:
        combined_matched = pd.DataFrame()

    if unmatched_dfs:
        combined_unmatched = pd.concat(unmatched_dfs, ignore_index=True)
    else:
        combined_unmatched = pd.DataFrame()

    os.makedirs(args.output_dir, exist_ok=True)
    combined_out_path = os.path.join(args.output_dir, "Unique_Commission_Import_Combined.csv")
    save_output_file(combined_matched, combined_out_path)
    print(" Done")

    elapsed_time = time.time() - start_time

    print("\n" + "=" * 76)
    print("📊 UNIQUE COMMISSION EXTRACTION SUMMARY")
    print("=" * 76)
    print(f" Unique Policy Keys Looked Up:  {len(valid_keys):,}")
    print(f" Commission File(s) Processed: {len(matched_dfs):,}")
    print(f" Total Unique Commission Rows:  {len(combined_matched):,}")
    print(f" Unmatched Commission Rows:     {len(combined_unmatched):,}")
    print(f" Consolidated Output File:      {os.path.abspath(combined_out_path)}")
    print(f" Unique Commission Folder:      {os.path.abspath(args.output_dir)}")
    if args.separate_unmatched and len(combined_unmatched) > 0:
        print(f" Unmatched Commission Folder:   {os.path.abspath(os.path.join(args.output_dir, 'unmatched_commission'))}")
    print(f" Elapsed Time:                  {elapsed_time:.1f}s")
    print("=" * 76)
    print("\n✅ Unique commission extraction completed successfully!\n")


if __name__ == "__main__":
    main()
