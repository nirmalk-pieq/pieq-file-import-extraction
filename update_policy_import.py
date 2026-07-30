#!/usr/bin/env python3
"""
Dynamic Policy List Enricher Script

Matches base Policy List Excel records with Master Policy records on Policy No and Agent Name,
populating Carrier, LOB, and Agent Level from the Master file into the output Excel file.
"""

import argparse
import os
import sys
import time
import pandas as pd


def clean_string_key(series: pd.Series) -> pd.Series:
    """Normalizes string keys (trim, collapse multi-spaces, uppercase)."""
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .apply(lambda s: " ".join(s.split()))
    )


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


def enrich_policy_data(
    base_file_path: str,
    master_file_path: str,
    output_file_path: str,
    base_policy_col: str = "Policy No",
    master_policy_col: str = "Policy Number",
    base_agent_col: str = "Agent Name",
    master_agent_col: str = "Agent Name",
    target_columns: list = None,
    ref_file_path: str = None,
    ref_key_col: str = "Agent Name",
    ref_val_col: str = None,
    ref_target_col: str = None,
) -> None:
    """Loads, matches, enriches, and saves policy data cleanly."""
    start_time = time.time()
    if target_columns is None:
        target_columns = ["Carrier", "LOB", "Agent Level", "Pay Mode", "Premium"]

    # Auto-append .xlsx extension if missing
    if not (output_file_path.endswith(".xlsx") or output_file_path.endswith(".xls")):
        output_file_path = f"{output_file_path}.xlsx"

    print("\n🚀 Starting Policy List Enrichment...\n")

    # 1. Validate files
    if not os.path.exists(base_file_path):
        raise FileNotFoundError(f"Base file not found: {base_file_path}")
    if not os.path.exists(master_file_path):
        raise FileNotFoundError(f"Master file not found: {master_file_path}")

    # 2. Load datasets
    print("[1/4] Loading input files...", end="", flush=True)
    base_df = pd.read_excel(base_file_path)
    master_df = pd.read_excel(master_file_path)

    # Normalize extra spaces in Agent Name
    if base_agent_col in base_df.columns:
        base_df[base_agent_col] = base_df[base_agent_col].fillna("").astype(str).apply(lambda s: " ".join(s.split()))
    if master_agent_col in master_df.columns:
        master_df[master_agent_col] = master_df[master_agent_col].fillna("").astype(str).apply(lambda s: " ".join(s.split()))

    print(f" Done (Base: {len(base_df):,} rows | Master: {len(master_df):,} rows)")

    # Validate columns
    for col in [base_policy_col, base_agent_col]:
        if col not in base_df.columns:
            raise KeyError(f"Column '{col}' missing in base file.")
    for col in [master_policy_col, master_agent_col]:
        if col not in master_df.columns:
            raise KeyError(f"Column '{col}' missing in master file.")

    available_target_cols = [col for col in target_columns if col in master_df.columns]
    missing_target_cols = [col for col in target_columns if col not in master_df.columns]
    if missing_target_cols:
        print(f"\n⚠️  Note: Column(s) {missing_target_cols} not present in master file, skipping them.")
    target_columns = available_target_cols

    # 3. Normalize & Match
    print("[2/4] Normalizing and matching policy records...", end="", flush=True)
    base_df["_pol_key"] = clean_string_key(base_df[base_policy_col])
    base_df["_agent_key"] = clean_string_key(base_df[base_agent_col])

    master_df["_pol_key"] = clean_string_key(master_df[master_policy_col])
    master_df["_agent_key"] = clean_string_key(master_df[master_agent_col])

    # Deduplicate master lookup map
    master_subset_cols = ["_pol_key", "_agent_key"] + target_columns
    master_subset = master_df[master_subset_cols].drop_duplicates(
        subset=["_pol_key", "_agent_key"], keep="first"
    )

    rename_dict = {col: f"{col}_master" for col in target_columns}
    master_target_cols = list(rename_dict.values())
    master_subset = master_subset.rename(columns=rename_dict)

    merged_df = pd.merge(base_df, master_subset, on=["_pol_key", "_agent_key"], how="left")
    print(" Done")

    # 4. Populate Target Columns
    print(f"[3/4] Populating {', '.join(target_columns)}...", end="", flush=True)
    matched_mask = merged_df[master_target_cols[0]].notna()
    matched_count = int(matched_mask.sum())
    total_count = len(base_df)
    unmatched_count = total_count - matched_count
    match_pct = (matched_count / total_count * 100) if total_count > 0 else 0.0

    for orig_col, m_col in zip(target_columns, master_target_cols):
        if orig_col in base_df.columns:
            merged_df[orig_col] = merged_df[m_col].combine_first(merged_df[orig_col])
        else:
            merged_df[orig_col] = merged_df[m_col]

    cleanup_cols = ["_pol_key", "_agent_key"] + master_target_cols
    output_df = merged_df.drop(columns=cleanup_cols)

    # Rename Pay Mode to Payment Frequency in output file
    if "Pay Mode" in output_df.columns:
        output_df = output_df.rename(columns={"Pay Mode": "Payment Frequency"})

    # Extract OVR rows (Product column contains 'ovr' case-insensitive)
    prod_col = None
    for col in output_df.columns:
        if str(col).strip().lower() in ("product", "product name"):
            prod_col = col
            break

    ovr_count = 0
    ovr_file_path = None
    if prod_col is not None:
        ovr_mask = output_df[prod_col].astype(str).str.contains("ovr", case=False, na=False)
        ovr_df = output_df[ovr_mask]
        output_df = output_df[~ovr_mask]
        matched_mask = matched_mask[~ovr_mask]
        ovr_count = len(ovr_df)

        if ovr_count > 0:
            if ref_file_path:
                ovr_df = apply_reference_updates(
                    df=ovr_df,
                    ref_file_path=ref_file_path,
                    key_col=ref_key_col,
                    val_col=ref_val_col,
                    target_col=ref_target_col,
                )
            output_dir = os.path.dirname(output_file_path)
            ovr_dir = os.path.join(output_dir, "ovr") if output_dir else "ovr"
            os.makedirs(ovr_dir, exist_ok=True)

            base_name = os.path.splitext(os.path.basename(output_file_path))[0]
            ovr_file_path = os.path.join(ovr_dir, f"{base_name}_OVR.xlsx")
            ovr_df = clean_id_and_quote_values(ovr_df)
            ovr_df.to_excel(ovr_file_path, index=False, engine='xlsxwriter')
            print(f"\n📦 Saved {ovr_count:,} OVR row(s) into '{ovr_file_path}'")

    print(" Done")

    # Apply reference updates to remaining output DataFrame
    if ref_file_path:
        output_df = apply_reference_updates(
            df=output_df,
            ref_file_path=ref_file_path,
            key_col=ref_key_col,
            val_col=ref_val_col,
            target_col=ref_target_col,
        )

    output_df = clean_id_and_quote_values(output_df)

    # 5. Export to Excel
    print("[4/4] Saving output file...", end="", flush=True)
    output_dir = os.path.dirname(output_file_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    output_df.to_excel(output_file_path, index=False, engine='xlsxwriter')
    print(" Done")

    elapsed_time = time.time() - start_time

    # Calculate matched stats
    matched_df = output_df[matched_mask]
    unmatched_df = output_df[~matched_mask]

    remaining_count = len(output_df)
    matched_count = len(matched_df)
    unmatched_count = len(unmatched_df)
    match_pct = (matched_count / remaining_count * 100) if remaining_count > 0 else 0.0

    unique_policies_matched = matched_df[base_policy_col].nunique()
    unique_policies_unmatched = unmatched_df[base_policy_col].nunique()

    # Crisp Summary Output
    print("\n" + "=" * 74)
    print("📊 EXECUTION SUMMARY")
    print("=" * 74)
    print(f" Total Base Rows:           {total_count:,}")
    if ovr_count > 0:
        print(f" OVR Rows Separated:        {ovr_count:,} ({ovr_file_path})")
        print(f" Remaining Output Rows:     {remaining_count:,}")
    print(f" Matched Records:           {matched_count:,} ({match_pct:.2f}%)")
    print(f" Unmatched Records:         {unmatched_count:,} ({100 - match_pct:.2f}%)")
    print(f" Unique Matched Policies:   {unique_policies_matched:,}")
    print(f" Output File:               {os.path.abspath(output_file_path)}")
    print(f" Elapsed Time:              {elapsed_time:.1f}s")
    print("=" * 74)

    # Prepare Complete Carrier & LOB Table including Matched and Unmatched
    display_df = output_df.copy()
    display_df["Carrier_Display"] = display_df["Carrier"].fillna("[UNMATCHED]")
    display_df["LOB_Display"] = display_df["LOB"].fillna("[UNMATCHED]")

    summary_table = display_df.groupby(["Carrier_Display", "LOB_Display"]).agg(
        Unique_Policies=(base_policy_col, "nunique"),
        Total_Records=(base_policy_col, "count")
    ).reset_index().sort_values(by="Total_Records", ascending=False)

    print("\n📋 COMPLETE CARRIER & LOB MATCH BREAKDOWN TABLE (ALL COMBINATIONS)")
    print("-" * 74)
    header = f"{'Carrier':<38} | {'LOB':<20} | {'Policies':<9} | {'Total':<7}"
    print(header)
    print("-" * 74)

    for _, row in summary_table.iterrows():
        c_name = str(row['Carrier_Display'])[:37]
        lob_name = str(row['LOB_Display'])[:19]
        print(f"{c_name:<38} | {lob_name:<20} | {row['Unique_Policies']:<9,} | {row['Total_Records']:<7,}")

    print("-" * 74)
    print(f"{'TOTAL MATCHED RECORDS':<61} | {unique_policies_matched:<9,} | {matched_count:<7,}")
    print(f"{'TOTAL UNMATCHED RECORDS':<61} | {unique_policies_unmatched:<9,} | {unmatched_count:<7,}")
    print("=" * 74)
    print(f"{'GRAND TOTAL (BASE DATASET)':<61} | {output_df[base_policy_col].nunique():<9,} | {total_count:<7,}")
    print("=" * 74)

    print("\n✅ Process completed successfully!\n")


def main():
    parser = argparse.ArgumentParser(
        description="Dynamic policy list enrichment tool."
    )
    parser.add_argument(
        "--base-file",
        type=str,
        required=True,
        help="Path to the base Policy List Excel file.",
    )
    parser.add_argument(
        "--master-file",
        type=str,
        required=True,
        help="Path to the Master Policy Excel file.",
    )
    parser.add_argument(
        "--output-file",
        type=str,
        required=True,
        help="Path where the populated output Excel file will be saved.",
    )
    parser.add_argument(
        "--base-policy-col",
        type=str,
        default="Policy No",
        help="Policy Number column in base file (default: 'Policy No').",
    )
    parser.add_argument(
        "--master-policy-col",
        type=str,
        default="Policy Number",
        help="Policy Number column in master file (default: 'Policy Number').",
    )
    parser.add_argument(
        "--base-agent-col",
        type=str,
        default="Agent Name",
        help="Agent Name column in base file (default: 'Agent Name').",
    )
    parser.add_argument(
        "--master-agent-col",
        type=str,
        default="Agent Name",
        help="Agent Name column in master file (default: 'Agent Name').",
    )
    parser.add_argument(
        "--target-cols",
        nargs="+",
        default=["Carrier", "LOB", "Agent Level", "Pay Mode", "Premium"],
        help="Target columns to pull from master file (default: Carrier LOB 'Agent Level' 'Pay Mode' Premium).",
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

    try:
        enrich_policy_data(
            base_file_path=args.base_file,
            master_file_path=args.master_file,
            output_file_path=args.output_file,
            base_policy_col=args.base_policy_col,
            master_policy_col=args.master_policy_col,
            base_agent_col=args.base_agent_col,
            master_agent_col=args.master_agent_col,
            target_columns=args.target_cols,
            ref_file_path=args.ref_file,
            ref_key_col=args.ref_key_col,
            ref_val_col=args.ref_val_col,
            ref_target_col=args.ref_target_col,
        )
    except Exception as e:
        print(f"\n❌ Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
