#!/usr/bin/env python3
"""
Policy List Import with Validation & Pipeline Orchestrator

Matches base Policy List Excel records with Master Policy records and Reference lookup files.
Populates target columns (Carrier, LOB, Agent Level, Agent ID, Pay Mode, Premium).
Identifies records with missing carriers or unmatched Agent IDs, extracts them into a
separate error Excel file with an 'Error' column, and outputs clean records.

Automatically proceeds to Step 2 (Commission Splits) and Step 3 (Carrier Splitting)
when zero errors are found, or prompts the user interactively when errors exist.
"""

import argparse
import os
import sys
import time
import subprocess
from typing import Tuple, List, Union
import pandas as pd
import numpy as np


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

    if not df_key_matched or not ref_key_matched:
        print(f" (WARNING: Key column '{key_col}' missing in dataset or reference file, skipping)")
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

        df[d_t_col] = mapped_series
        updated_col_names.append(d_t_col)

    ref_keys_set = set(ref_clean["_norm_key"])
    df["_ref_agent_matched"] = df_norm_key.isin(ref_keys_set)

    print(f" Done ({total_updated:,} record(s) updated in {updated_col_names})")
    return df


def clean_id_and_quote_values(df: pd.DataFrame) -> pd.DataFrame:
    """Strips leading/trailing quotes from text columns and removes '.0' suffix from numeric IDs."""
    df = df.loc[:, ~df.columns.duplicated()].copy()

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
        target_series = df[col]
        
        # If still a DataFrame due to edge cases, pick first column
        if isinstance(target_series, pd.DataFrame):
            target_series = target_series.iloc[:, 0]
            
        if c_lower == "to month":
            df[col] = target_series.apply(
                lambda x: "" if pd.isna(x) or str(x).strip().lstrip("'").rstrip("'").strip('"') in ("0", "0.0") else str(x).strip().lstrip("'").rstrip("'").strip('"')
            )
        elif col in id_cols or "id" in c_lower or "no" in c_lower or "number" in c_lower:
            df[col] = target_series.apply(_clean_val)
        elif target_series.dtype == object or target_series.dtype == 'string':
            df[col] = target_series.apply(lambda x: str(x).lstrip("'").rstrip("'").strip('"') if pd.notna(x) else x)

    return df


def is_value_missing(val) -> bool:
    """Checks if a cell value is missing, empty, or an unmatched placeholder (including '\\N')."""
    if pd.isna(val):
        return True
    s = str(val).strip()
    if not s or s.upper() in ("", "NAN", "NONE", "NULL", "[UNMATCHED]", "UNMATCHED", r"\N", r"\\N"):
        return True
    return False


def propagate_carrier_by_policy(df: pd.DataFrame) -> pd.DataFrame:
    """
    Groups DataFrame by (Policy No, Product) and propagates non-empty policy-level metadata
    (Carrier, LOB, Payment Frequency, Premium) across all rows belonging to the same policy.
    """
    if df is None or len(df) == 0:
        return df

    pol_col = None
    prod_col = None
    for col in df.columns:
        c_lower = str(col).strip().lower()
        if c_lower in ("policy no", "policy number", "policyno", "policynumber") and not pol_col:
            pol_col = col
        elif c_lower in ("product", "product name", "prod") and not prod_col:
            prod_col = col

    if not pol_col:
        return df

    meta_cols = ["Carrier", "LOB", "Pay Mode", "Payment Frequency", "Premium"]
    target_meta_cols = [c for c in meta_cols if c in df.columns]

    if not target_meta_cols:
        return df

    pol_key_ser = clean_string_key(df[pol_col])
    prod_key_ser = clean_string_key(df[prod_col]) if prod_col else pd.Series([""] * len(df), index=df.index)

    temp_df = pd.DataFrame({"_g_pol": pol_key_ser, "_g_prod": prod_key_ser}, index=df.index)

    def _fill_series(series):
        valid = series.dropna()
        valid = valid[valid.astype(str).str.strip().str.upper().replace({"\\N": "", "NAN": "", "NONE": ""}) != ""]
        if not valid.empty:
            first_val = valid.iloc[0]
            return series.apply(
                lambda x: first_val if pd.isna(x) or str(x).strip().upper() in ("", "\\N", "NAN", "NONE") else x
            )
        return series

    group_cols = ["_g_pol", "_g_prod"] if prod_col else ["_g_pol"]

    for m_col in target_meta_cols:
        temp_df[m_col] = df[m_col]
        df[m_col] = temp_df.groupby(group_cols, sort=False)[m_col].transform(_fill_series)

    return df


def separate_errors(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Evaluates dataset records for missing Carrier, missing/unmatched Agent ID, and master policy matching.
    Returns (clean_df, error_df) where error_df contains an added 'Error' column.
    """
    error_list = []
    
    carrier_col = None
    agent_id_col = None
    has_master_matched_flag = "_master_matched" in df.columns
    has_ref_agent_matched_flag = "_ref_agent_matched" in df.columns

    for col in df.columns:
        c_lower = str(col).strip().lower()
        if c_lower == "carrier":
            carrier_col = col
        elif c_lower == "agent id":
            agent_id_col = col

    for _, row in df.iterrows():
        reasons = []

        # Check Master File match
        if has_master_matched_flag and not bool(row.get("_master_matched", True)):
            reasons.append("Policy Not found")

        # Check Carrier
        if carrier_col is None or is_value_missing(row.get(carrier_col)):
            reasons.append("Carrier is missing")

        # Check Agent ID (missing or Agent Name not matched in reference file)
        is_agent_id_err = (
            agent_id_col is None
            or is_value_missing(row.get(agent_id_col))
            or (has_ref_agent_matched_flag and not bool(row.get("_ref_agent_matched", True)))
        )
        if is_agent_id_err:
            reasons.append("Agent ID not matched")

        if reasons:
            error_list.append(", ".join(reasons))
        else:
            error_list.append(None)

    df_copy = df.copy()
    df_copy["_error_reason"] = error_list

    error_mask = df_copy["_error_reason"].notna()
    
    clean_df = df_copy[~error_mask].drop(columns=["_error_reason"])
    error_df = df_copy[error_mask].copy()
    error_df["Error"] = error_df["_error_reason"]
    error_df = error_df.drop(columns=["_error_reason"])

    for temp_col in ["_master_matched", "_ref_agent_matched"]:
        if temp_col in clean_df.columns:
            clean_df = clean_df.drop(columns=[temp_col])
        if temp_col in error_df.columns:
            error_df = error_df.drop(columns=[temp_col])

    return clean_df, error_df


def enrich_policy_data_with_validation(
    base_file_path: str,
    master_file_path: str,
    output_file_path: str,
    error_file_path: str = None,
    fixed_error_file_path: str = None,
    base_policy_col: str = "Policy No",
    master_policy_col: str = "Policy Number",
    base_agent_col: str = "Agent Name",
    master_agent_col: str = "Agent Name",
    target_columns: list = None,
    ref_file_path: str = None,
    ref_key_col: str = "Agent Name",
    ref_val_col: str = None,
    ref_target_col: str = None,
) -> int:
    """
    Loads, matches, enriches policy records, filters out missing carriers / agent IDs into an error Excel file,
    and exports clean records. Returns total error count.
    """
    start_time = time.time()
    if target_columns is None:
        target_columns = ["Carrier", "LOB", "Agent Level", "Agent ID", "Pay Mode", "Premium"]

    _, out_ext = os.path.splitext(output_file_path)
    if os.path.isdir(output_file_path) or output_file_path.endswith(("/", "\\")) or not out_ext:
        out_dir = output_file_path.rstrip("/\\")
        if not out_dir:
            out_dir = "."
        os.makedirs(out_dir, exist_ok=True)
        output_file_path = os.path.join(out_dir, "output.xlsx")

    out_dir = os.path.dirname(output_file_path) or "."
    os.makedirs(out_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(output_file_path))[0]
    if not base_name or base_name in (".xlsx", ".xls"):
        base_name = "output"

    if not error_file_path:
        error_file_path = os.path.join(out_dir, f"{base_name}_Errors.xlsx") if out_dir else f"{base_name}_Errors.xlsx"

    if not (error_file_path.endswith(".xlsx") or error_file_path.endswith(".xls")):
        error_file_path = f"{error_file_path}.xlsx"

    print("\n🚀 Starting Policy List Enrichment & Validation...\n")

    # 1. Validate files
    if not os.path.exists(base_file_path):
        raise FileNotFoundError(f"Base file not found: {base_file_path}")
    if not os.path.exists(master_file_path):
        raise FileNotFoundError(f"Master file not found: {master_file_path}")

    # 2. Load datasets
    print("[1/5] Loading input files...", end="", flush=True)
    master_df = pd.read_excel(master_file_path)

    if fixed_error_file_path and os.path.exists(fixed_error_file_path):
        fixed_err_df = pd.read_excel(fixed_error_file_path)
        if "Error" in fixed_err_df.columns:
            fixed_err_df = fixed_err_df.drop(columns=["Error"])

        if os.path.exists(output_file_path):
            clean_part = pd.read_excel(output_file_path)
            base_df = pd.concat([clean_part, fixed_err_df], ignore_index=True)
            print(f" Done (Merged clean output {len(clean_part):,} rows + fixed error file {len(fixed_err_df):,} rows = {len(base_df):,} total rows | Master: {len(master_df):,} rows)")
        else:
            raw_base = pd.read_excel(base_file_path)
            b_keys = clean_string_key(raw_base[base_policy_col]) + "||" + clean_string_key(raw_base[base_agent_col])
            f_keys = clean_string_key(fixed_err_df[base_policy_col]) + "||" + clean_string_key(fixed_err_df[base_agent_col])
            raw_base_clean = raw_base[~b_keys.isin(set(f_keys))]
            base_df = pd.concat([raw_base_clean, fixed_err_df], ignore_index=True)
            print(f" Done (Updated base file with {len(fixed_err_df):,} fixed error rows = {len(base_df):,} total rows | Master: {len(master_df):,} rows)")
    else:
        base_df = pd.read_excel(base_file_path)
        print(f" Done (Base: {len(base_df):,} rows | Master: {len(master_df):,} rows)")

    base_df = base_df.loc[:, ~base_df.columns.duplicated()].copy()
    master_df = master_df.loc[:, ~master_df.columns.duplicated()].copy()

    # Validate required matching columns
    for col in [base_policy_col, base_agent_col]:
        if col not in base_df.columns:
            raise KeyError(f"Column '{col}' missing in base file.")
    for col in [master_policy_col, master_agent_col]:
        if col not in master_df.columns:
            raise KeyError(f"Column '{col}' missing in master file.")

    available_target_cols = [col for col in target_columns if col in master_df.columns]
    missing_target_cols = [col for col in target_columns if col not in master_df.columns]
    if missing_target_cols:
        print(f"\n⚠️ Note: Target column(s) {missing_target_cols} not present in master file, skipping lookup for them.")
    target_columns = available_target_cols

    # 3. Normalize & Match
    print("[2/5] Normalizing and matching policy records...", end="", flush=True)
    base_df["_pol_key"] = clean_string_key(base_df[base_policy_col])
    base_df["_agent_key"] = clean_string_key(base_df[base_agent_col])

    master_df["_pol_key"] = clean_string_key(master_df[master_policy_col])
    master_df["_agent_key"] = clean_string_key(master_df[master_agent_col])

    master_subset_cols = ["_pol_key", "_agent_key"] + target_columns
    master_subset = master_df[master_subset_cols].drop_duplicates(
        subset=["_pol_key", "_agent_key"], keep="first"
    )

    rename_dict = {col: f"{col}_master" for col in target_columns}
    master_target_cols = list(rename_dict.values())
    master_subset = master_subset.rename(columns=rename_dict)
    master_subset["_master_matched"] = True

    merged_df = pd.merge(base_df, master_subset, on=["_pol_key", "_agent_key"], how="left")
    merged_df["_master_matched"] = merged_df["_master_matched"].fillna(False).astype(bool)
    print(" Done")

    # 4. Populate Target Columns
    print(f"[3/5] Populating {', '.join(target_columns)}...", end="", flush=True)
    for orig_col, m_col in zip(target_columns, master_target_cols):
        if orig_col in base_df.columns:
            m_ser = merged_df[m_col].astype(object)
            b_ser = merged_df[orig_col].astype(object)
            
            m_missing = m_ser.apply(is_value_missing)
            
            # Where master is missing (\N, NaN, empty), preserve base value; otherwise use master value
            merged_df[orig_col] = pd.Series(np.where(m_missing, b_ser, m_ser), index=merged_df.index)
        else:
            merged_df[orig_col] = merged_df[m_col]

    cleanup_cols = ["_pol_key", "_agent_key"] + master_target_cols
    output_df = merged_df.drop(columns=cleanup_cols)

    if "Pay Mode" in output_df.columns:
        output_df = output_df.rename(columns={"Pay Mode": "Payment Frequency"})

    # Extract OVR product rows if present
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
            ovr_dir = os.path.join(output_dir, "policy_ovr") if output_dir else "policy_ovr"
            os.makedirs(ovr_dir, exist_ok=True)

            base_name = os.path.splitext(os.path.basename(output_file_path))[0]
            ovr_file_path = os.path.join(ovr_dir, f"{base_name}_policy_OVR.xlsx")
            ovr_df = clean_id_and_quote_values(ovr_df)
            ovr_df.to_excel(ovr_file_path, index=False)
            print(f"\n📦 Saved {ovr_count:,} policy OVR row(s) into '{ovr_file_path}'")

    print(" Done")

    # Apply reference updates if provided
    if ref_file_path:
        output_df = apply_reference_updates(
            df=output_df,
            ref_file_path=ref_file_path,
            key_col=ref_key_col,
            val_col=ref_val_col,
            target_col=ref_target_col,
        )

    output_df = clean_id_and_quote_values(output_df)
    output_df = propagate_carrier_by_policy(output_df)

    # 5. Separate Error Rows
    print("[4/5] Separating missing carrier & unmatched agent ID errors...", end="", flush=True)
    clean_df, error_df = separate_errors(output_df)
    error_count = len(error_df)
    clean_count = len(clean_df)
    total_processed = clean_count + error_count
    print(f" Done ({clean_count:,} clean records | {error_count:,} error records)")

    # 6. Save Output Files
    print("[5/5] Exporting clean Excel file and error Excel file...", end="", flush=True)
    output_dir = os.path.dirname(output_file_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    error_dir = os.path.dirname(error_file_path)
    if error_dir and not os.path.exists(error_dir):
        os.makedirs(error_dir, exist_ok=True)

    clean_df.to_excel(output_file_path, index=False)
    error_df.to_excel(error_file_path, index=False)
    print(" Done")

    elapsed_time = time.time() - start_time

    # Summary Output
    print("\n" + "=" * 74)
    print("📊 EXECUTION SUMMARY (ENRICHMENT & ERROR SEPARATION)")
    print("=" * 74)
    print(f" Total Base Rows Processed:  {len(base_df):,}")
    if ovr_count > 0:
        print(f" OVR Rows Separated:         {ovr_count:,} ({ovr_file_path})")
    print(f" Total Evaluated Rows:       {total_processed:,}")
    print(f" ✅ Clean Records Exported:   {clean_count:,} ({output_file_path})")
    print(f" ⚠️ Error Records Exported:   {error_count:,} ({error_file_path})")
    print(f" Elapsed Time:               {elapsed_time:.1f}s")
    print("=" * 74)

    if error_count > 0:
        print("\n📋 ERROR REASON BREAKDOWN TABLE")
        print("-" * 74)
        header = f"{'Error Reason':<50} | {'Record Count':<18}"
        print(header)
        print("-" * 74)
        err_summary = error_df["Error"].value_counts().reset_index()
        err_summary.columns = ["Error Reason", "Record Count"]
        for _, row in err_summary.iterrows():
            print(f"{str(row['Error Reason']):<50} | {row['Record Count']:<18,}")
        print("-" * 74)
        print(f"\n⚠️ Total Error Count: {error_count:,}")
        print(f"⚠️ Records with missing/unmatched data were separated into: {error_file_path}\n")

    print("\n✅ Step 1 (Policy List Enrichment & Validation) completed successfully!\n")
    return error_count, output_file_path, error_file_path


def run_cmd(cmd_args: list):
    """Runs a subprocess command and prints output, exiting on failure."""
    print(f"\nCommand: {' '.join(cmd_args)}")
    result = subprocess.run(cmd_args, capture_output=False, text=True)
    if result.returncode != 0:
        print(f"\n❌ Error: Command failed with exit code {result.returncode}", file=sys.stderr)
        sys.exit(result.returncode)


def combine_clean_and_fixed_errors(clean_file_path: str, fixed_error_file_path: str, combined_output_path: str) -> str:
    """Joins clean policy output file and fixed error file into a single combined master file."""
    print(f"\n🔄 Combining clean output ('{os.path.basename(clean_file_path)}') and fixed error file ('{os.path.basename(fixed_error_file_path)}')...")
    clean_df = pd.read_excel(clean_file_path)
    fixed_err_df = pd.read_excel(fixed_error_file_path)

    if "Error" in fixed_err_df.columns:
        fixed_err_df = fixed_err_df.drop(columns=["Error"])

    combined_df = pd.concat([clean_df, fixed_err_df], ignore_index=True)
    
    out_dir = os.path.dirname(combined_output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        
    combined_df.to_excel(combined_output_path, index=False)
    print(f"✅ Combined {len(clean_df):,} clean row(s) + {len(fixed_err_df):,} fixed error row(s) = {len(combined_df):,} total row(s) into '{combined_output_path}'.")
    return combined_output_path


def handle_interactive_error_selection(
    clean_file_path: str,
    error_file_path: str,
    output_dir: str,
    error_count: int,
    auto_confirm: bool = False
) -> Tuple[bool, str]:
    """
    Displays an interactive breakdown of error reasons and prompts the user to select
    which error categories (e.g. 1, 2, 'all', 'none') to include in the output dataset for Step 2 & 3.

    Returns:
        Tuple[should_proceed: bool, lookup_file_path: str]
    """
    if error_count <= 0 or not os.path.exists(error_file_path):
        return True, clean_file_path

    try:
        error_df = pd.read_excel(error_file_path)
    except Exception as e:
        print(f"⚠️ Could not read error file '{error_file_path}': {e}")
        return False, clean_file_path

    if error_df.empty or "Error" not in error_df.columns:
        return True, clean_file_path

    err_summary = error_df["Error"].value_counts().reset_index()
    err_summary.columns = ["Error Reason", "Record Count"]
    options_list = list(err_summary.itertuples(index=False))

    combined_output_path = os.path.join(output_dir, "Policy_List_Combined.xlsx")

    def _combine_all():
        print(f"\n🔄 Combining clean output and ALL error records...")
        clean_df = pd.read_excel(clean_file_path)
        err_clean_df = error_df.drop(columns=["Error"], errors="ignore")
        combined_df = pd.concat([clean_df, err_clean_df], ignore_index=True)
        os.makedirs(output_dir, exist_ok=True)
        combined_df.to_excel(combined_output_path, index=False)
        print(f"✅ Combined {len(clean_df):,} clean row(s) + {len(err_clean_df):,} error row(s) = {len(combined_df):,} total row(s) into '{combined_output_path}'.")
        return True, combined_output_path

    if auto_confirm:
        print(f"\n⚡ Auto-confirm (--yes) enabled. Combining all error records and proceeding to Step 2 & Step 3...")
        return _combine_all()

    if not sys.stdin.isatty():
        print(f"\nℹ️ Non-interactive session detected. Combining all error records and proceeding to Step 2 & Step 3...")
        return _combine_all()

    print(f"\n⚠️ Errors were detected ({error_count:,} record(s)).\n")
    print("📋 ERROR SELECTION OPTIONS FOR PIPELINE STEP 2 & STEP 3:")
    print("-" * 74)
    print(f"{'Option':<8} | {'Error Reason':<48} | {'Record Count':<12}")
    print("-" * 74)
    for idx, row in enumerate(options_list, 1):
        reason, count = row[0], row[1]
        print(f"  [{idx}]   | {str(reason):<48} | {count:<12,}")
    print("-" * 74)
    print("Options:")
    print("  • Enter option numbers (e.g., '1,2' or '1') to include specific error categories & proceed")
    print("  • Enter 'all' (or 'y') to include ALL error categories & proceed")
    print("  • Enter 'none' (or '0') to exclude ALL error categories (use clean records only) & proceed")
    print("  • Enter 'n' (or 'exit') to pause pipeline without running Step 2 & 3")

    while True:
        try:
            resp = input("\n❓ Select option(s) [all / none / 1,2,... / n]: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\n")
            return False, clean_file_path

        if resp in ("n", "no", "exit", "cancel"):
            return False, clean_file_path

        if resp in ("all", "y", "yes"):
            return _combine_all()

        if resp in ("none", "0"):
            print(f"\nℹ️ Excluding all error records. Proceeding to Step 2 & 3 with clean records only ({os.path.basename(clean_file_path)})...")
            return True, clean_file_path

        tokens = [t.strip() for t in resp.replace(",", " ").split() if t.strip()]
        valid_indices = []
        is_valid = True
        for t in tokens:
            if t.isdigit():
                val = int(t)
                if 1 <= val <= len(options_list):
                    valid_indices.append(val)
                else:
                    is_valid = False
                    break
            else:
                is_valid = False
                break

        if is_valid and valid_indices:
            selected_indices = list(dict.fromkeys(valid_indices))
            selected_reasons = [options_list[i - 1][0] for i in selected_indices]

            selected_mask = error_df["Error"].isin(selected_reasons)
            selected_err_df = error_df[selected_mask].drop(columns=["Error"], errors="ignore")
            excluded_err_df = error_df[~selected_mask]

            clean_df = pd.read_excel(clean_file_path)
            combined_df = pd.concat([clean_df, selected_err_df], ignore_index=True)

            os.makedirs(output_dir, exist_ok=True)
            combined_df.to_excel(combined_output_path, index=False)

            opt_str = ", ".join(f"[{i}]" for i in selected_indices)
            print(f"\n✅ Selected option(s) {opt_str}: Included {len(selected_err_df):,} error row(s) into '{combined_output_path}'.")
            print(f"✅ Total dataset for Step 2 & Step 3: {len(clean_df):,} clean row(s) + {len(selected_err_df):,} selected error row(s) = {len(combined_df):,} total row(s).")
            if len(excluded_err_df) > 0:
                print(f"⚠️ Excluded {len(excluded_err_df):,} error row(s) (remaining in '{os.path.basename(error_file_path)}').")

            return True, combined_output_path

        print(f"❌ Invalid selection '{resp}'. Please enter option numbers (e.g. '1, 2'), 'all', 'none', or 'n'.")


def run_pipeline_steps_2_and_3(
    clean_output_file: str,
    base_split_path: str,
    output_dir: str = None,
    split_col: str = "Carrier",
    no_splits: bool = False,
    ref_file_path: str = None,
    ref_key_col: str = "Agent Name",
    ref_val_col: str = None,
    ref_target_col: str = None,
):
    """Executes Step 2 (Commission Split & Agent Level) and Step 3 (Carrier Splitting)."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if not output_dir:
        output_dir = os.path.dirname(clean_output_file) or "."

    comm_import_dir = os.path.join(output_dir, "commission_import")
    split_dir = os.path.join(output_dir, "split_policy")

    # --- Step 2: Commission Split Calculation & Agent Level Enrichment ---
    print("\n--- [STEP 2/3] Commission Split Calculation & Agent Level Enrichment ---")
    enrich_script = os.path.join(script_dir, "enrich_policy_commission_import.py")
    step2_args = [
        sys.executable,
        enrich_script,
        "--base-file", base_split_path,
        "--master-file", clean_output_file,
        "--output-dir", comm_import_dir
    ]
    if not os.path.isdir(base_split_path):
        base_name = os.path.splitext(os.path.basename(base_split_path))[0]
        step2_out = os.path.join(comm_import_dir, f"{base_name}_enriched.csv")
        step2_args.extend(["--output-file", step2_out])

    if ref_file_path:
        step2_args.extend(["--ref-file", ref_file_path])
        if ref_key_col:
            step2_args.extend(["--ref-key-col", ref_key_col])
        if ref_val_col:
            step2_args.extend(["--ref-val-col", ref_val_col])
        if ref_target_col:
            step2_args.extend(["--ref-target-col", ref_target_col])

    run_cmd(step2_args)

    # --- Step 3: Carrier Splitting ---
    if not no_splits:
        print("\n--- [STEP 3/3] Carrier Splitting (Clean Policy Import File) ---")
        split_script = os.path.join(script_dir, "split_file.py")
        step3_args = [
            sys.executable,
            split_script,
            "--file", clean_output_file,
            "--column", split_col,
            "--output-dir", split_dir,
            "--format", "match"
        ]
        run_cmd(step3_args)
    else:
        print("\n--- [STEP 3/3] Carrier Splitting: Skipped ---")

    print("\n==========================================================================")
    print("✅ ENTIRE PIPELINE COMPLETED SUCCESSFULLY")
    print("==========================================================================")


def main():
    parser = argparse.ArgumentParser(
        description="Policy list import with validation, error separation, and pipeline orchestrator."
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
        help="Path where the clean populated output Excel file will be saved.",
    )
    parser.add_argument(
        "--error-file",
        type=str,
        default=None,
        help="Path where the error records Excel file will be saved (default: output_file_name_Errors.xlsx).",
    )
    parser.add_argument(
        "--fixed-error-file",
        type=str,
        default=None,
        help="Optional path to resolved/fixed error file to combine with clean output before Step 2.",
    )
    parser.add_argument(
        "--base-policy-col",
        type=str,
        default="Policy No",
        help="Policy Number column name in base file (default: 'Policy No').",
    )
    parser.add_argument(
        "--master-policy-col",
        type=str,
        default="Policy Number",
        help="Policy Number column name in master file (default: 'Policy Number').",
    )
    parser.add_argument(
        "--base-agent-col",
        type=str,
        default="Agent Name",
        help="Agent Name column name in base file (default: 'Agent Name').",
    )
    parser.add_argument(
        "--master-agent-col",
        type=str,
        default="Agent Name",
        help="Agent Name column name in master file (default: 'Agent Name').",
    )
    parser.add_argument(
        "--target-cols",
        nargs="+",
        default=["Carrier", "LOB", "Agent Level", "Agent ID", "Pay Mode", "Premium"],
        help="Target columns to pull from master file.",
    )
    parser.add_argument(
        "--ref-file",
        type=str,
        default=None,
        help="Optional reference/mapping file (.xlsx/.csv) to update output columns.",
    )
    parser.add_argument(
        "--ref-key-col",
        type=str,
        default="Agent Name",
        help="Matching key column in reference file (default: 'Agent Name').",
    )
    parser.add_argument(
        "--ref-val-col",
        type=str,
        default=None,
        help="Source value column in reference file.",
    )
    parser.add_argument(
        "--ref-target-col",
        type=str,
        default=None,
        help="Target column in output file to update.",
    )
    parser.add_argument(
        "--base-split",
        "--commission-file",
        type=str,
        default=None,
        help="Path to the base APL Commission Split CSV file/folder for Step 2.",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="Path to directory for saving Step 2 and Step 3 outputs.",
    )
    parser.add_argument(
        "--split-col",
        type=str,
        default="Carrier",
        help="Column to split the Step 1 output file by (default: 'Carrier').",
    )
    parser.add_argument(
        "--no-splits",
        action="store_true",
        help="Skip Step 3 carrier splitting.",
    )
    parser.add_argument(
        "--yes",
        "-y",
        "--auto-confirm",
        action="store_true",
        help="Auto-confirm prompt and proceed to next steps without stopping.",
    )

    args = parser.parse_args()

    try:
        error_count, clean_out_path, err_out_path = enrich_policy_data_with_validation(
            base_file_path=args.base_file,
            master_file_path=args.master_file,
            output_file_path=args.output_file,
            error_file_path=args.error_file,
            fixed_error_file_path=args.fixed_error_file,
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

        if args.base_split:
            step2_lookup_file = clean_out_path
            out_dir = os.path.dirname(clean_out_path) or "."

            if args.fixed_error_file and os.path.exists(args.fixed_error_file):
                combined_path = os.path.join(out_dir, "Policy_List_Combined.xlsx")
                step2_lookup_file = combine_clean_and_fixed_errors(
                    clean_file_path=step2_lookup_file,
                    fixed_error_file_path=args.fixed_error_file,
                    combined_output_path=combined_path
                )

            if error_count == 0:
                print("\n✨ Zero errors found in Step 1! Automatically proceeding to next pipeline steps (Step 2 & Step 3)...")
                should_proceed = True
            else:
                should_proceed, step2_lookup_file = handle_interactive_error_selection(
                    clean_file_path=step2_lookup_file,
                    error_file_path=err_out_path,
                    output_dir=out_dir,
                    error_count=error_count,
                    auto_confirm=args.yes,
                )

            if should_proceed:
                run_pipeline_steps_2_and_3(
                    clean_output_file=step2_lookup_file,
                    base_split_path=args.base_split,
                    output_dir=args.output_dir,
                    split_col=args.split_col,
                    no_splits=args.no_splits,
                    ref_file_path=args.ref_file,
                    ref_key_col=args.ref_key_col,
                    ref_val_col=args.ref_val_col,
                    ref_target_col=args.ref_target_col,
                )
            else:
                print(f"\n🛑 Pipeline paused. Please resolve errors in the error file, then re-run with --fixed-error-file <path_to_fixed_errors>.\n")

    except Exception as e:
        print(f"\n❌ Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

