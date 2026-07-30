#!/usr/bin/env python3
"""
Policy Sequence Adjustment & Enrichment Script

Processes a base CSV and a master Excel file, performs sequential amount adjustments,
populates Agent Level via a composite lookup, and saves the final result.
"""

import argparse
import os
import sys
import time
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


def parse_amount(val) -> float:
    """Parses currency strings into float values."""
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace("$", "").replace(",", "").strip()
    if not val_str:
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0


def format_amount(val: float) -> str:
    """Formats float values back to standard currency strings."""
    if val < 0:
        return f"-${abs(val):,.2f}"
    return f"${val:,.2f}"


def clean_id_and_quote_values(df: pd.DataFrame) -> pd.DataFrame:
    """Strips leading/trailing quotes from text columns and removes '.0' suffix from numeric IDs."""
    def _clean_val(val):
        if pd.isna(val):
            return val
        s = str(val).strip()
        s = s.lstrip("'").rstrip("'").strip('"')
        if s.endswith(".0"):
            s = s[:-2]
        return s.strip()

    id_cols = ["Policy No", "Agent ID", "EACID", "Policy Number"]
    for col in df.columns:
        c_lower = str(col).strip().lower()
        if c_lower == "to month":
            df[col] = df[col].apply(
                lambda x: "" if pd.isna(x) or str(x).strip().lstrip("'").rstrip("'").strip('"') in ("0", "0.0") else str(x).strip().lstrip("'").rstrip("'").strip('"')
            )
        elif col in id_cols or "id" in c_lower or "no" in c_lower or "number" in c_lower:
            df[col] = df[col].apply(_clean_val)
        elif df[col].dtype == object or df[col].dtype == 'string':
            df[col] = df[col].apply(lambda x: str(x).lstrip("'").rstrip("'").strip('"') if pd.notna(x) else x)

    return df


def apply_reference_updates(
    df: pd.DataFrame,
    ref_file_path: str,
    key_col: str = "Agent Name",
    val_col: str = None,
    target_col: str = None
) -> pd.DataFrame:
    """
    Dynamically replaces/updates column values in df using a reference mapping file.
    Matches records on key_col (case & whitespace insensitive).
    If val_col is specified, updates target_col (defaults to val_col).
    If val_col is None, automatically updates all matching non-key columns present in ref_file.
    """
    if not ref_file_path or not os.path.exists(ref_file_path):
        return df

    print(f"\n📋 Applying reference update from '{ref_file_path}'...", end="", flush=True)

    try:
        _, ext = os.path.splitext(ref_file_path.lower())
        if ext in ('.xlsx', '.xls'):
            ref_df = pd.read_excel(ref_file_path)
        elif ext == '.csv':
            ref_df = pd.read_csv(ref_file_path)
        else:
            print(f" (WARNING: Unsupported reference file extension '{ext}', skipping)")
            return df
    except Exception as e:
        print(f" (WARNING: Failed to load reference file: {e}, skipping)")
        return df

    df_key_matched = None
    for c in df.columns:
        if str(c).strip().lower() == str(key_col).strip().lower():
            df_key_matched = c
            break

    ref_key_matched = None
    for c in ref_df.columns:
        if str(c).strip().lower() == str(key_col).strip().lower():
            ref_key_matched = c
            break

    if not df_key_matched:
        print(f" (WARNING: Key column '{key_col}' not found in dataset, skipping)")
        return df

    if not ref_key_matched:
        print(f" (WARNING: Key column '{key_col}' not found in reference file, skipping)")
        return df

    def clean_key(s):
        return (
            s.fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
            .apply(lambda x: " ".join(x.split()))
            .replace({"nan": "", "none": ""})
        )

    ref_clean = ref_df.dropna(subset=[ref_key_matched]).copy()
    ref_clean["_norm_key"] = clean_key(ref_clean[ref_key_matched])
    ref_clean = ref_clean[ref_clean["_norm_key"] != ""].drop_duplicates(subset=["_norm_key"], keep="first")

    df_norm_key = clean_key(df[df_key_matched])

    cols_to_update = []
    if val_col:
        t_col = target_col if target_col else val_col
        ref_v_col = None
        for c in ref_df.columns:
            if str(c).strip().lower() == str(val_col).strip().lower():
                ref_v_col = c
                break
        if ref_v_col:
            cols_to_update.append((ref_v_col, t_col))
        else:
            print(f" (WARNING: Value column '{val_col}' not found in reference file, skipping)")
            return df
    else:
        for r_col in ref_df.columns:
            if str(r_col).strip().lower() == str(ref_key_matched).strip().lower():
                continue
            for d_col in df.columns:
                if str(r_col).strip().lower() == str(d_col).strip().lower():
                    cols_to_update.append((r_col, d_col))
                    break

    if not cols_to_update:
        print(" (WARNING: No matching value columns found to update, skipping)")
        return df

    total_updated = 0
    updated_col_names = []

    for r_v_col, d_t_col in cols_to_update:
        mapping_dict = dict(zip(ref_clean["_norm_key"], ref_clean[r_v_col]))
        mapped_series = df_norm_key.map(mapping_dict)

        updated_count = int(mapped_series.notna().sum())
        total_updated = max(total_updated, updated_count)

        if d_t_col in df.columns:
            df[d_t_col] = mapped_series.combine_first(df[d_t_col])
        else:
            df[d_t_col] = mapped_series
        updated_col_names.append(d_t_col)

    print(f" Done ({total_updated:,} record(s) updated in {updated_col_names})")
    return df


def process_single_split_file(
    base_file_path: str,
    master_df: pd.DataFrame,
    ref_file_path: str = None,
    ref_key_col: str = "Agent Name",
    ref_val_col: str = None,
    ref_target_col: str = None,
) -> pd.DataFrame:
    """Processes sequence adjustments, populates master lookup, and applies reference updates for a single split file."""
    _, ext = os.path.splitext(base_file_path.lower())
    if ext in ('.xlsx', '.xls'):
        base_df = pd.read_excel(base_file_path)
    else:
        base_df = pd.read_csv(base_file_path)

    if "Agent Name" in base_df.columns:
        base_df["Agent Name"] = base_df["Agent Name"].fillna("").astype(str).apply(lambda s: " ".join(s.split()))

    required_base_cols = ["Policy No", "Product", "Agent Name", "Sequence", "From Month", "Amount"]
    for col in required_base_cols:
        if col not in base_df.columns:
            raise KeyError(f"Required column '{col}' missing in '{os.path.basename(base_file_path)}'.")

    # Sequence adjustments
    base_df["_amt_numeric"] = base_df["Amount"].apply(parse_amount)
    group_seq_amts = {}
    for idx, row in base_df.iterrows():
        pol = row["Policy No"]
        fm = row["From Month"]
        seq = int(pd.to_numeric(row["Sequence"], errors="coerce"))
        amt = row["_amt_numeric"]

        key = (pol, fm)
        if key not in group_seq_amts:
            group_seq_amts[key] = {}
        if seq not in group_seq_amts[key]:
            group_seq_amts[key][seq] = amt

    new_amounts = []
    for idx, row in base_df.iterrows():
        pol = row["Policy No"]
        fm = row["From Month"]
        seq = int(pd.to_numeric(row["Sequence"], errors="coerce"))
        amt = row["_amt_numeric"]
        orig_amt_str = row["Amount"]

        if amt > 0:
            prev_seq = seq - 1
            key = (pol, fm)
            if key in group_seq_amts and prev_seq in group_seq_amts[key]:
                prev_amt = group_seq_amts[key][prev_seq]
                new_amt_num = amt - prev_amt
                new_amounts.append(format_amount(new_amt_num))
            else:
                new_amounts.append(orig_amt_str)
        else:
            new_amounts.append(orig_amt_str)

    base_df["Amount"] = new_amounts
    base_df = base_df.drop(columns=["_amt_numeric"])

    # Master Lookup Mapping
    base_df["_pol_key"] = clean_key(base_df["Policy No"])
    base_df["_prod_key"] = clean_key(base_df["Product"])
    base_df["_agent_key"] = clean_key(base_df["Agent Name"])

    master_df["_pol_key"] = clean_key(master_df["Policy No"])
    master_df["_prod_key"] = clean_key(master_df["Product"])
    master_df["_agent_key"] = clean_key(master_df["Agent Name"])

    lookup_cols = ["_pol_key", "_prod_key", "_agent_key"]
    for col in ["Agent Level", "Carrier", "LOB", "Pay Mode", "Payment Frequency", "Premium"]:
        if col in master_df.columns:
            lookup_cols.append(col)

    master_lookup = master_df[lookup_cols].drop_duplicates(
        subset=["_pol_key", "_prod_key", "_agent_key"], keep="first"
    )

    merged_df = pd.merge(
        base_df,
        master_lookup,
        on=["_pol_key", "_prod_key", "_agent_key"],
        how="left"
    )

    output_df = merged_df.drop(columns=["_pol_key", "_prod_key", "_agent_key"])

    if ref_file_path:
        output_df = apply_reference_updates(
            df=output_df,
            ref_file_path=ref_file_path,
            key_col=ref_key_col,
            val_col=ref_val_col,
            target_col=ref_target_col,
        )

    output_df = clean_id_and_quote_values(output_df)
    return output_df


def save_carrier_splits(df: pd.DataFrame, split_dir: str):
    """Splits enriched commission DataFrame by Carrier column and writes CSV files into split_dir."""
    if "Carrier" not in df.columns or len(df) == 0:
        return

    os.makedirs(split_dir, exist_ok=True)
    carrier_series = df["Carrier"].fillna("unassigned_Carrier")

    split_count = 0
    for carrier_name, group_df in df.groupby(carrier_series):
        safe_name = str(carrier_name).strip()
        safe_name = "".join(c for c in safe_name if c.isalnum() or c in (" ", "_", "-")).strip()
        if not safe_name:
            safe_name = "unassigned_Carrier"

        out_file = os.path.join(split_dir, f"{safe_name}.csv")
        group_df.to_csv(out_file, index=False)
        split_count += 1

    print(f"\n📦 Carrier splits for Commission Import saved in '{split_dir}': ({split_count} carrier file(s) created)")


def main():
    parser = argparse.ArgumentParser(
        description="Process base CSV file/folder and master Excel file for policy sequence adjustment and carrier split."
    )
    parser.add_argument(
        "--base-file",
        "--base-split",
        dest="base_file",
        type=str,
        required=True,
        help="Path to the base CSV/Excel file OR directory containing base split files.",
    )
    parser.add_argument(
        "--master-file",
        type=str,
        required=True,
        help="Path to the master Excel file.",
    )
    parser.add_argument(
        "--output-file",
        type=str,
        default=None,
        help="Path to save output CSV file (if single file input).",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="Path to directory where commission_import and carrier splits will be saved.",
    )
    parser.add_argument(
        "--ref-file",
        "--reference-file",
        "--update-file",
        "--mapping-file",
        "--agent-roster",
        "--agent-roster-file",
        "--roster-file",
        dest="ref_file",
        type=str,
        default=None,
        help="Optional reference/mapping file (Excel/CSV) to dynamically update dataset column values.",
    )
    parser.add_argument(
        "--ref-key-col",
        "--key-col",
        "--agent-roster-key-col",
        "--roster-key-col",
        dest="ref_key_col",
        type=str,
        default="Agent Name",
        help="Matching key column in reference file (default: 'Agent Name').",
    )
    parser.add_argument(
        "--ref-val-col",
        "--val-col",
        "--agent-roster-val-col",
        "--roster-val-col",
        dest="ref_val_col",
        type=str,
        default=None,
        help="Source value column in reference file to pull values from (default: auto-discovers matching columns).",
    )
    parser.add_argument(
        "--ref-target-col",
        "--target-col",
        "--agent-roster-target-col",
        "--roster-target-col",
        dest="ref_target_col",
        type=str,
        default=None,
        help="Target column in output file to update (default: same as ref-val-col).",
    )

    args = parser.parse_args()

    start_time = time.time()
    print("🚀 Starting Policy List Processing & Enrichment for bulk policy import...")

    # 1. Validate inputs
    if not os.path.exists(args.base_file):
        print(f"❌ Error: Base file/folder not found at {args.base_file}", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(args.master_file):
        print(f"❌ Error: Master file not found at {args.master_file}", file=sys.stderr)
        sys.exit(1)

    # 2. Determine output directories
    if args.output_dir:
        base_out_dir = args.output_dir
        if base_out_dir.rstrip("/").endswith("commission_import"):
            comm_import_dir = base_out_dir
        else:
            comm_import_dir = os.path.join(base_out_dir, "commission_import")
    elif args.output_file:
        comm_import_dir = os.path.dirname(args.output_file) or "commission_import"
    else:
        comm_import_dir = "commission_import"

    split_comm_dir = os.path.join(comm_import_dir, "split_commission")
    os.makedirs(comm_import_dir, exist_ok=True)

    # 3. Load Master File
    print("[1/3] Loading master Excel file...", end="", flush=True)
    try:
        master_df = pd.read_excel(args.master_file)
        if "Agent Name" in master_df.columns:
            master_df["Agent Name"] = master_df["Agent Name"].fillna("").astype(str).apply(lambda s: " ".join(s.split()))
        print(f" Done ({len(master_df):,} rows)")
    except Exception as e:
        print(f"\n❌ Error loading master file: {e}", file=sys.stderr)
        sys.exit(1)

    required_master_cols = ["Policy No", "Product", "Agent Name", "Agent Level"]
    for col in required_master_cols:
        if col not in master_df.columns:
            print(f"❌ Error: Required column '{col}' missing in master file.", file=sys.stderr)
            sys.exit(1)

    # 4. Collect input files to process
    if os.path.isdir(args.base_file):
        input_files = [
            os.path.join(args.base_file, f) for f in sorted(os.listdir(args.base_file))
            if f.lower().endswith(('.csv', '.xlsx', '.xls')) and not f.startswith('~$')
        ]
        if not input_files:
            print(f"❌ Error: No CSV/Excel files found in folder '{args.base_file}'.", file=sys.stderr)
            sys.exit(1)
        print(f"[2/3] Processing {len(input_files)} file(s) from directory '{args.base_file}'...")
    else:
        input_files = [args.base_file]
        print(f"[2/3] Processing single file '{args.base_file}'...")

    processed_dfs = []
    total_processed_rows = 0

    for idx, filepath in enumerate(input_files, 1):
        filename = os.path.basename(filepath)
        print(f"  ({idx}/{len(input_files)}) Enriching '{filename}'...", end="", flush=True)
        try:
            out_df = process_single_split_file(
                base_file_path=filepath,
                master_df=master_df,
                ref_file_path=args.ref_file,
                ref_key_col=args.ref_key_col,
                ref_val_col=args.ref_val_col,
                ref_target_col=args.ref_target_col,
            )
            processed_dfs.append(out_df)
            total_processed_rows += len(out_df)
            print(f" Done ({len(out_df):,} rows)")

            # Save individual enriched file into comm_import_dir
            if len(input_files) == 1 and args.output_file:
                single_out_path = args.output_file
            else:
                raw_name = os.path.splitext(filename)[0]
                single_out_path = os.path.join(comm_import_dir, f"{raw_name}_enriched.csv")

            os.makedirs(os.path.dirname(single_out_path), exist_ok=True)
            out_df.to_csv(single_out_path, index=False)
            print(f"     💾 Saved: {single_out_path}")

        except Exception as e:
            print(f" FAILED ({e})")
            continue

    if not processed_dfs:
        print("\n❌ Error: No files were successfully processed.", file=sys.stderr)
        sys.exit(1)

    combined_df = pd.concat(processed_dfs, ignore_index=True)

    # 5. Split commission import data by Carrier inside commission_import/split_commission/
    print("[3/3] Generating Carrier-wise splits for Commission Import...", end="", flush=True)
    save_carrier_splits(combined_df, split_comm_dir)
    print(" Done")

    elapsed_time = time.time() - start_time
    matched_count = combined_df["Agent Level"].notna().sum() if "Agent Level" in combined_df.columns else 0
    match_pct = (matched_count / total_processed_rows * 100) if total_processed_rows > 0 else 0.0

    print("\n" + "=" * 74)
    print("📊 EXECUTION SUMMARY")
    print("=" * 74)
    print(f" Input Files Processed:     {len(processed_dfs):,}")
    print(f" Total Rows Processed:      {total_processed_rows:,}")
    print(f" Matched Agent Levels:      {matched_count:,} ({match_pct:.2f}%)")
    print(f" Commission Import Folder:  {os.path.abspath(comm_import_dir)}")
    print(f" Carrier Split Directory:   {os.path.abspath(split_comm_dir)}")
    print(f" Elapsed Time:              {elapsed_time:.1f}s")
    print("=" * 74)
    print("\n✅ Process completed successfully!\n")


if __name__ == "__main__":
    main()
