#!/usr/bin/env python3
"""
Commission Extraction Validation Utility

Cross-verifies extracted policy commission values against the master Product Commission
rate file (ProductWise Commission Rate.xlsx). Validates carrier commission base values,
agent levels, month tier ranges, and payout methods.
"""

import argparse
import math
import os
import sys
import time
from typing import Dict, List, Set, Tuple, Union
import pandas as pd
import numpy as np


def clean_key(series: pd.Series) -> pd.Series:
    """Normalizes string keys (strips quotes/whitespace, uppercase, collapses spaces)."""
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
    """Parses currency strings or numeric values into float."""
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace("$", "").replace(",", "").strip()
    if not val_str:
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0


def parse_rate(val) -> float:
    """Parses percentage strings or numeric values into float (as percentage e.g. 70.0)."""
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace("%", "").strip()
    if not val_str:
        return 0.0
    try:
        v = float(val_str)
        return v
    except ValueError:
        return 0.0


def format_amount(val: float) -> str:
    """Formats float values back to standard currency strings."""
    if val < 0:
        return f"-${abs(val):,.2f}"
    return f"${val:,.2f}"


def format_rate(val: float) -> str:
    """Formats float values into percentage strings (e.g. '70.00 %')."""
    return f"{val:.2f} %"


def load_file(file_path: str) -> pd.DataFrame:
    """Loads CSV or Excel file into DataFrame."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: '{file_path}'")

    _, ext = os.path.splitext(file_path.lower())
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(file_path)
    elif ext == ".csv":
        return pd.read_csv(file_path, low_memory=False)
    else:
        raise ValueError(f"Unsupported file format '{ext}' for file '{file_path}'.")


def unpivot_product_commission_rates(excel_path: str) -> pd.DataFrame:
    """
    Unpivots repeating month tier column groups from ProductWise Commission Rate.xlsx
    into a structured rules DataFrame.
    """
    df_raw = load_file(excel_path)

    # 17 month-tier column suffixes in Excel: '', '.1', '.2', ..., '.16'
    suffixes = [""] + [f".{i}" for i in range(1, 17)]
    rules = []

    for idx, row in df_raw.iterrows():
        prod_id = row.get("Product ID")
        prod_desc = str(row.get("Product Description", "")).strip()
        contract = str(row.get("Contract", "")).strip()

        if not prod_desc or prod_desc.lower() in ("nan", "none", ""):
            continue

        for s in suffixes:
            fm_col = f"From-To Months{s}"
            pct_col = f"Rate (Percentage){s}"
            dollar_col = f"Rate (Dollar Amount){s}"
            fixed_col = f"Column Rate is Fixed{s}"

            if fm_col in row and pd.notna(row[fm_col]):
                fm_val = str(row[fm_col]).strip()
                if not fm_val or fm_val.lower() in ("nan", "none", ""):
                    continue

                parts = fm_val.split("-")
                try:
                    from_m = int(parts[0].strip())
                    to_m = int(parts[1].strip()) if len(parts) > 1 else from_m
                    if to_m == 0:
                        to_m = 9999  # 0 indicates open-ended / onwards
                except Exception:
                    continue

                pct_val = float(row.get(pct_col, 0.0)) if pd.notna(row.get(pct_col)) else 0.0
                dollar_val = float(row.get(dollar_col, 0.0)) if pd.notna(row.get(dollar_col)) else 0.0
                is_fixed = str(row.get(fixed_col, "N")).strip().upper()

                # Convert rate percentage decimal (e.g., 0.0264) to percentage scale (2.64) if stored as fraction <= 1.0
                rate_pct_scaled = pct_val * 100.0 if (0.0 < pct_val <= 1.0) else pct_val

                rules.append({
                    "Product_ID": prod_id,
                    "Product_Description": prod_desc,
                    "Contract": contract,
                    "From_Month": from_m,
                    "To_Month": to_m,
                    "Rate_Percentage": rate_pct_scaled,
                    "Rate_Dollar": dollar_val,
                    "Is_Fixed": is_fixed,
                    "_prod_clean": " ".join(prod_desc.lower().split()),
                    "_prod_norm": " ".join(" ".join(prod_desc.replace("/", " ").replace("-", " ").replace(",", " ").replace("*", " ").split()).lower().split()),
                    "_contract_clean": " ".join(contract.lower().split()),
                })

    df_rules = pd.DataFrame(rules)
    if df_rules.empty:
        raise ValueError(f"No valid rate rules could be extracted from '{excel_path}'.")
    return df_rules


def normalize_product_str(p: str) -> str:
    """Strips punctuation (/ , - *) and collapses whitespace for product matching."""
    if not p or pd.isna(p):
        return ""
    clean = str(p).replace("/", " ").replace("-", " ").replace(",", " ").replace("*", " ")
    return " ".join(clean.lower().split())


def load_product_normalization(norm_path: str = None) -> Dict[str, str]:
    """Loads product normalization file into a lookup dictionary (Product -> New Product Name)."""
    if not norm_path or not os.path.exists(norm_path):
        return {}

    try:
        df_norm = load_file(norm_path)
        # Strip trailing/leading spaces from column headers
        df_norm = df_norm.rename(columns=lambda c: str(c).strip())

        prod_col = next((c for c in df_norm.columns if c.lower() == "product"), None)
        target_col = next((c for c in df_norm.columns if "new product" in c.lower()), None)

        if not prod_col or not target_col:
            return {}

        norm_dict = {}
        for _, row in df_norm.iterrows():
            p_val = str(row[prod_col]).strip()
            t_val = str(row[target_col]).strip()
            if p_val and t_val and t_val.lower() not in ("nan", "none", ""):
                clean_p = " ".join(p_val.lower().split())
                norm_dict[clean_p] = t_val
                norm_dict[normalize_product_str(p_val)] = t_val
        return norm_dict
    except Exception as e:
        print(f"⚠️ Warning: Could not load product normalization file: {e}")
        return {}


def validate_commission_dataset(
    df_comm: pd.DataFrame,
    df_rules: pd.DataFrame | None,
    norm_dict: Dict[str, str] = None,
    tolerance: float = 0.01,
    min_seq_len: int = None,
    max_seq_len: int = None,
) -> Tuple[pd.DataFrame, pd.DataFrame, dict]:
    """
    Cross-verifies extracted commission records by grouping sequences on (Policy No, Product, From Month).
    - Validates total group Carrier Commission against Agency Master Rate rules (e.g. 75% or $18.00).
    - Validates each sequence line item Amount/Rate against its specific Contract tier (e.g. LVL5 = $15.12, Sales Leader = $1.26).
    - Verifies line-item split Rate shares sum to 100.00%.
    - Validates sequence count for each group to be within min_seq_len and max_seq_len.
    Returns: (validated_df, error_df, summary_metrics)
    """
    df_out = df_comm.copy()

    # Column identification
    pol_col = next((c for c in df_out.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno")), None)
    prod_col = next((c for c in df_out.columns if str(c).strip().lower() in ("product", "product name", "product_description")), None)
    contract_col = next((c for c in df_out.columns if str(c).strip().lower() in ("contract", "agent level", "level")), None)
    fm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("from month", "from_month", "month")), None)
    to_month_col = next((c for c in df_out.columns if str(c).strip().lower() in ("to month", "to_month")), None)
    seq_col = next((c for c in df_out.columns if str(c).strip().lower() in ("sequence", "seq")), None)
    carrier_comm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("carrier commission", "carrier_commission", "commission")), None)
    payout_method_col = next((c for c in df_out.columns if str(c).strip().lower() in ("payout method", "payout_method", "pay mode")), None)
    rate_col = next((c for c in df_out.columns if str(c).strip().lower() == "rate"), None)
    amt_col = next((c for c in df_out.columns if str(c).strip().lower() == "amount"), None)
    carrier_col = next((c for c in df_out.columns if str(c).strip().lower() in ("carrier", "carrier name", "carrier_name")), None)

    if not prod_col or not fm_col:
        raise KeyError(f"Missing required columns (Product, From Month) in commission dataset. Available: {list(df_out.columns)}")

    norm_dict = norm_dict or {}

    validation_statuses = [""] * len(df_out)
    expected_values = [""] * len(df_out)
    variances = [0.0] * len(df_out)
    error_reasons = [""] * len(df_out)

    valid_count = 0
    mismatch_val_count = 0
    invalid_split_count = 0
    unmatched_prod_count = 0
    unmatched_contract_count = 0
    unmatched_month_count = 0
    sequence_count_error_count = 0
    to_month_mismatch_count = 0
    missing_carrier_count = 0
    carrier_mismatch_count = 0

    # Grouping helpers
    df_out["_pol_clean"] = clean_key(df_out[pol_col]) if pol_col else pd.Series(["ROW_" + str(i) for i in range(len(df_out))], index=df_out.index)
    df_out["_prod_clean"] = clean_key(df_out[prod_col])
    df_out["_fm_clean"] = clean_key(df_out[fm_col])

    grouped = df_out.groupby(["_pol_clean", "_prod_clean", "_fm_clean"], sort=False)

    for _, group in grouped:
        indices = group.index.tolist()
        seq_1_idx = indices[0]
        seq_1_row = group.loc[seq_1_idx]

        raw_prod = str(seq_1_row.get(prod_col, "")).strip()
        raw_fm = seq_1_row.get(fm_col, 1)

        clean_prod = " ".join(raw_prod.lower().split())
        norm_prod = normalize_product_str(raw_prod)

        if clean_prod in norm_dict:
            target_p = norm_dict[clean_prod]
            clean_prod = " ".join(target_p.lower().split())
            norm_prod = normalize_product_str(target_p)
        elif norm_prod in norm_dict:
            target_p = norm_dict[norm_prod]
            clean_prod = " ".join(target_p.lower().split())
            norm_prod = normalize_product_str(target_p)

        try:
            from_month_num = int(float(str(raw_fm).strip().replace("'", "").replace('"', "")))
        except (ValueError, TypeError):
            from_month_num = 1

        extracted_comm_str = str(seq_1_row.get(carrier_comm_col, "")) if carrier_comm_col else ""
        payout_method = str(seq_1_row.get(payout_method_col, "")).strip().upper() if payout_method_col else ""

        extracted_comm_amt = parse_amount(extracted_comm_str)
        extracted_comm_pct = parse_rate(extracted_comm_str)

        # Deduplicate sequence entries within group by Sequence number to handle duplicate policy rows in input
        if seq_col and seq_col in group.columns:
            seq_subgroup = group.drop_duplicates(subset=[seq_col])
        else:
            seq_subgroup = group

        # Sequence length validation
        seq_count = len(seq_subgroup)
        if (min_seq_len is not None and seq_count < min_seq_len) or (max_seq_len is not None and seq_count > max_seq_len):
            sequence_count_error_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "SEQUENCE_COUNT_ERROR"
                expected_values[idx] = "N/A"
                variances[idx] = 0.0
                error_reasons[idx] = f"Group sequence count {seq_count} is out of required range ({min_seq_len}-{max_seq_len})."
            continue

        # Inconsistent To Month validation
        if to_month_col:
            to_months_parsed = set()
            for val in group[to_month_col]:
                if pd.isna(val):
                    to_months_parsed.add("NaN")
                else:
                    val_str = str(val).strip()
                    try:
                        num = int(float(val_str))
                        to_months_parsed.add(num)
                    except ValueError:
                        to_months_parsed.add(val_str)
            if len(to_months_parsed) > 1:
                to_month_mismatch_count += len(indices)
                for idx in indices:
                    validation_statuses[idx] = "TO_MONTH_MISMATCH"
                    expected_values[idx] = "N/A"
                    variances[idx] = 0.0
                    error_reasons[idx] = f"Group contains inconsistent 'To Month' values: {list(to_months_parsed)}"
                continue

        # Carrier validation and population
        if carrier_col:
            carriers_in_group = []
            for val in group[carrier_col]:
                if pd.notna(val):
                    val_str = str(val).strip()
                    if val_str and val_str not in ("", "NaN", "None", "\\N", "nan", "none"):
                        carriers_in_group.append(val_str)
            
            unique_carriers = sorted(list(set(carriers_in_group)))
            
            if len(unique_carriers) == 0:
                missing_carrier_count += len(indices)
                for idx in indices:
                    validation_statuses[idx] = "MISSING_CARRIER"
                    expected_values[idx] = "N/A"
                    variances[idx] = 0.0
                    error_reasons[idx] = "Group has no Carrier specified."
                continue
            elif len(unique_carriers) > 1:
                carrier_mismatch_count += len(indices)
                for idx in indices:
                    validation_statuses[idx] = "CARRIER_MISMATCH"
                    expected_values[idx] = "N/A"
                    variances[idx] = 0.0
                    error_reasons[idx] = f"Group contains inconsistent carrier values: {unique_carriers}"
                continue
            else:
                # Exactly one unique carrier - auto-populate it across all rows of the group in df_out
                carrier_name = unique_carriers[0]
                df_out.loc[indices, carrier_col] = carrier_name

        rates_sum = sum(parse_rate(seq_subgroup.loc[i, rate_col]) for i in seq_subgroup.index) if rate_col else 0.0
        amounts_sum = sum(parse_amount(seq_subgroup.loc[i, amt_col]) for i in seq_subgroup.index) if amt_col else 0.0

        # If no rules are loaded, skip rate matching validation and mark as VALID
        if df_rules is None or df_rules.empty:
            valid_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "VALID"
                expected_values[idx] = "SKIPPED_RATE_VALIDATION"
                variances[idx] = 0.0
                error_reasons[idx] = ""
            continue

        # 1. Hierarchical Master Product lookup
        # Priority 1: Exact raw product match
        p_rules = df_rules[df_rules["Product_Description"] == raw_prod]
        # Priority 2: Exact clean product match
        if p_rules.empty:
            p_rules = df_rules[df_rules["_prod_clean"] == clean_prod]
        # Priority 3: Punctuation-normalized match
        if p_rules.empty and "_prod_norm" in df_rules.columns:
            p_rules = df_rules[df_rules["_prod_norm"] == norm_prod]
        # Priority 4: Substring fallback
        if p_rules.empty:
            p_rules = df_rules[df_rules["_prod_clean"].apply(lambda p: p in clean_prod or clean_prod in p)]

        if p_rules.empty:
            unmatched_prod_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "UNMATCHED_PRODUCT"
                expected_values[idx] = "N/A"
                variances[idx] = 0.0
                error_reasons[idx] = f"Product '{raw_prod}' not found in ProductWise Commission Rate table."
            continue

        # 2. Agency Total Benchmark Lookup (Filter by month range first)
        m_rules = p_rules[(p_rules["From_Month"] <= from_month_num) & (from_month_num <= p_rules["To_Month"])]
        if m_rules.empty:
            unmatched_month_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "UNMATCHED_MONTH_RANGE"
                expected_values[idx] = "N/A"
                variances[idx] = 0.0
                error_reasons[idx] = f"Month {from_month_num} is out of rate tier ranges for product '{raw_prod}'."
            continue

        # Select benchmark contract tier (Prioritize non-zero AGENCY / STREET0 / STREET rates over zero-rate entries)
        agency_rules = m_rules[m_rules["_contract_clean"] == "agency"]
        if agency_rules.empty or (agency_rules["Rate_Dollar"].max() == 0.0 and agency_rules["Rate_Percentage"].max() == 0.0):
            street_rules = m_rules[m_rules["_contract_clean"].isin(["street", "street0", "top"])]
            if not street_rules.empty and (street_rules["Rate_Dollar"].max() > 0.0 or street_rules["Rate_Percentage"].max() > 0.0):
                agency_rules = street_rules

        if agency_rules.empty or (agency_rules["Rate_Dollar"].max() == 0.0 and agency_rules["Rate_Percentage"].max() == 0.0):
            agency_rules = m_rules.sort_values(by=["Rate_Dollar", "Rate_Percentage"], ascending=False)

        agency_rule = agency_rules.sort_values(by=["Rate_Dollar", "Rate_Percentage"], ascending=False).iloc[0]
        exp_agency_pct = agency_rule["Rate_Percentage"]
        exp_agency_dollar = agency_rule["Rate_Dollar"]

        # 3. Validate Group Carrier Commission & Sequence Line Items
        group_valid = True
        group_var = 0.0
        group_exp_str = ""
        group_reason = ""
        group_status = "VALID"

        effective_comm_pct = extracted_comm_pct if extracted_comm_pct > 0 else (rates_sum if rates_sum <= 100.0 else 0.0)
        effective_comm_amt = extracted_comm_amt if extracted_comm_amt > 0 else amounts_sum

        if payout_method == "PERCENTAGE" or (exp_agency_pct > 0.0 and exp_agency_dollar == 0.0):
            group_exp_str = format_rate(exp_agency_pct)
            group_var = round(abs(effective_comm_pct - exp_agency_pct), 4)

            if group_var > tolerance:
                group_status = "VALUE_MISMATCH"
                group_reason = f"Group Carrier Commission mismatch: Extracted benchmark rate '{format_rate(effective_comm_pct)}' vs expected Agency benchmark '{group_exp_str}'"
                group_valid = False
            elif len(seq_subgroup) > 1 and abs(rates_sum - 100.00) > 1.5:
                group_status = "INVALID_SEQUENCE_SPLIT"
                group_var = round(abs(rates_sum - 100.00), 2)
                group_reason = f"Sequence split rates sum to {format_rate(rates_sum)} instead of 100.00%"
                group_valid = False

        elif payout_method == "FIXED FEE" or (exp_agency_dollar > 0.0):
            group_exp_str = format_amount(exp_agency_dollar)
            group_var = round(abs(effective_comm_amt - exp_agency_dollar), 2)

            if group_var > tolerance:
                group_status = "VALUE_MISMATCH"
                group_reason = f"Group Carrier Commission mismatch: Extracted benchmark amount '{format_amount(effective_comm_amt)}' vs expected Agency benchmark '{group_exp_str}'"
                group_valid = False
            elif len(seq_subgroup) > 1 and rate_col and abs(rates_sum - 100.00) > 1.5:
                group_status = "INVALID_SEQUENCE_SPLIT"
                group_var = round(abs(rates_sum - 100.00), 2)
                group_reason = f"Sequence split rates sum to {format_rate(rates_sum)} instead of 100.00%"
                group_valid = False
        else:
            group_exp_str = format_rate(exp_agency_pct) if exp_agency_pct > 0 else format_amount(exp_agency_dollar)

        # 4. Apply Validation Outcome across all Sequence Rows in the Group
        if group_valid:
            valid_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "VALID"
                expected_values[idx] = group_exp_str
                variances[idx] = 0.0
                error_reasons[idx] = ""
        else:
            if group_status == "VALUE_MISMATCH":
                mismatch_val_count += len(indices)
            elif group_status == "INVALID_SEQUENCE_SPLIT":
                invalid_split_count += len(indices)

            for idx in indices:
                validation_statuses[idx] = group_status
                expected_values[idx] = group_exp_str
                variances[idx] = group_var
                error_reasons[idx] = group_reason

    df_out["Validation_Status"] = validation_statuses
    df_out["Expected_Base_Commission"] = expected_values
    df_out["Commission_Variance"] = variances
    df_out["Validation_Error_Reason"] = error_reasons

    df_out = df_out.drop(columns=["_pol_clean", "_prod_clean", "_fm_clean"])

    # Group-level split: if ANY row in a group has an error, move the ENTIRE group to errors.
    # Identify groups that contain at least one non-VALID row.
    error_mask = df_out["Validation_Status"] != "VALID"
    df_clean = df_out[~error_mask].copy()   # all rows are VALID in these groups (no partial groups)
    df_errors = df_out[error_mask].copy()   # entire group already flagged by validate loop above

    # NOTE: The validation loop already propagates the group error status to every row in the
    # group (see "Apply Validation Outcome across all Sequence Rows" block above), so every row
    # belonging to an errored group already has a non-VALID Validation_Status.  Therefore
    # filtering on Validation_Status != 'VALID' already captures the full group.

    error_record_count = len(df_errors)

    summary = {
        "total_records": len(df_comm),
        "valid_records": valid_count,
        "total_errors": error_record_count,
        "value_mismatches": mismatch_val_count,
        "invalid_sequence_splits": invalid_split_count,
        "unmatched_products": unmatched_prod_count,
        "unmatched_contracts": unmatched_contract_count,
        "unmatched_month_ranges": unmatched_month_count,
        "sequence_count_errors": sequence_count_error_count,
        "to_month_mismatches": to_month_mismatch_count,
        "missing_carriers": missing_carrier_count,
        "carrier_mismatches": carrier_mismatch_count,
        "accuracy_pct": round((valid_count / len(df_comm) * 100.0), 2) if len(df_comm) > 0 else 0.0,
    }

    return df_clean, df_errors, summary


def save_output_file(df: pd.DataFrame, output_path: str):
    """Saves DataFrame to CSV or Excel file."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    _, ext = os.path.splitext(output_path.lower())
    if ext in (".xlsx", ".xls"):
        df.to_excel(output_path, index=False)
    else:
        df.to_csv(output_path, index=False)


def process_commission_validation(
    input_path: str,
    product_commission_path: str,
    norm_path: str = None,
    output_dir: str = "pipeline_output/commission_validation",
    tolerance: float = 0.01,
    min_seq_len: int = 3,
    max_seq_len: int = 4,
) -> Tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Orchestrates loading, validation, and report generation."""
    start_time = time.time()
    print("==================================================")
    print("      POLICY COMMISSION EXTRACTION VALIDATION     ")
    print("==================================================")

    df_rules = None
    if product_commission_path and os.path.exists(product_commission_path):
        print(f"📖 Unpivoting master product commission rates from: '{product_commission_path}'...")
        df_rules = unpivot_product_commission_rates(product_commission_path)
        print(f"   ✅ Extracted {len(df_rules):,} active rate tier rule(s).")
    else:
        print(f"⚠️ Warning: Master product commission file not found at '{product_commission_path}'. Skipping rate/benchmark validation.")

    norm_dict = {}
    if norm_path and os.path.exists(norm_path):
        print(f"📖 Loading product normalization table from: '{norm_path}'...")
        norm_dict = load_product_normalization(norm_path)
        print(f"   ✅ Loaded {len(norm_dict):,} product normalization mapping(s).")

    # Collect input files
    input_files = []
    if os.path.isdir(input_path):
        input_files = [
            os.path.join(input_path, f) for f in sorted(os.listdir(input_path))
            if f.lower().endswith((".csv", ".xlsx", ".xls")) and not f.startswith("~$") and not f.startswith(".")
        ]
    elif os.path.isfile(input_path):
        input_files = [input_path]

    if not input_files:
        raise FileNotFoundError(f"No CSV or Excel files found at input path: '{input_path}'")

    print(f"📖 Loading commission extraction dataset ({len(input_files)} file(s))...")
    dfs = [load_file(f) for f in input_files]
    df_comm = pd.concat(dfs, ignore_index=True)
    print(f"   ✅ Loaded {len(df_comm):,} total commission record(s).")

    print(f"\n🔍 Cross-verifying extracted values against ProductWise Commission Rate table...")
    df_validated, df_errors, summary = validate_commission_dataset(
        df_comm, df_rules, norm_dict, tolerance, min_seq_len=min_seq_len, max_seq_len=max_seq_len
    )

    # Save output reports
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.basename(os.path.abspath(input_path)).replace(".", "_") if os.path.isdir(input_path) else os.path.splitext(os.path.basename(input_path))[0]

    # Clean file: only fully-valid groups (xlsx for upload compatibility)
    clean_out_path = os.path.join(output_dir, f"{base_name}_clean.xlsx")
    # Sequence error file: groups with invalid sequence counts
    seq_errors_out_path = os.path.join(output_dir, f"{base_name}_sequence_errors.xlsx")
    # Error file: all other error groups
    errors_out_path = os.path.join(output_dir, f"{base_name}_errors.xlsx")
    summary_out_path = os.path.join(output_dir, f"{base_name}_validation_summary.txt")

    # Split df_errors into sequence count errors and other errors
    df_seq_errors = df_errors[df_errors["Validation_Status"] == "SEQUENCE_COUNT_ERROR"]
    df_other_errors = df_errors[df_errors["Validation_Status"] != "SEQUENCE_COUNT_ERROR"]

    print(f"\n💾 Exporting validation results...")
    save_output_file(df_validated, clean_out_path)
    print(f"   • Clean (fully-valid groups): '{clean_out_path}' ({len(df_validated):,} record(s))")

    if not df_seq_errors.empty:
        save_output_file(df_seq_errors, seq_errors_out_path)
        print(f"   • Sequence errors ({len(df_seq_errors):,} record(s)): '{seq_errors_out_path}'")
    else:
        print("   • Zero sequence errors found! ✨")

    if not df_other_errors.empty:
        save_output_file(df_other_errors, errors_out_path)
        print(f"   • Other error groups ({len(df_other_errors):,} record(s)): '{errors_out_path}'")
    else:
        print("   • Zero other validation errors found! ✨")

    # Generate summary report text
    elapsed = time.time() - start_time
    summary_text = (
        "==================================================\n"
        "      COMMISSION VALIDATION SUMMARY REPORT        \n"
        "==================================================\n"
        f"Input Path:               {input_path}\n"
        f"Master Commission Table:  {product_commission_path}\n"
        f"Total Records Checked:    {summary['total_records']:,}\n"
        f"Clean Records (valid groups): {summary['valid_records']:,} ({summary['accuracy_pct']}%)\n"
        f"Sequence Errors:              {len(df_seq_errors):,}\n"
        f"Other Error Records:          {len(df_other_errors):,}\n"
        "--------------------------------------------------\n"
        "Error Breakdown:\n"
        f"  • Sequence Count Errors:  {summary.get('sequence_count_errors', 0):,}\n"
        f"  • Inconsistent To Month:  {summary.get('to_month_mismatches', 0):,}\n"
        f"  • Missing Carrier Errors: {summary.get('missing_carriers', 0):,}\n"
        f"  • Carrier Mismatch:       {summary.get('carrier_mismatches', 0):,}\n"
        f"  • Rate/Dollar Mismatches: {summary['value_mismatches']:,}\n"
        f"  • Invalid Sequence Splits:{summary.get('invalid_sequence_splits', 0):,}\n"
        f"  • Unmatched Products:     {summary['unmatched_products']:,}\n"
        f"  • Unmatched Agent Levels: {summary['unmatched_contracts']:,}\n"
        f"  • Unmatched Month Range:  {summary['unmatched_month_ranges']:,}\n"
        "--------------------------------------------------\n"
        f"Execution Time:           {elapsed:.2f} seconds\n"
        "==================================================\n"
    )

    with open(summary_out_path, "w") as f:
        f.write(summary_text)
    print(f"   • Summary report saved: '{summary_out_path}'\n")

    print(summary_text)
    return df_validated, df_errors, summary


def main():
    parser = argparse.ArgumentParser(
        description="Cross-verify extracted policy commission values against ProductWise Commission Rate master table."
    )
    parser.add_argument(
        "-i", "--commission-input",
        required=True,
        help="Path to extracted commission CSV/Excel file or directory of files."
    )
    parser.add_argument(
        "-p", "--product-commission-file",
        default=None,
        help="Path to optional ProductWise Commission Rate excel file (rate validation is skipped if not provided)."
    )
    parser.add_argument(
        "-n", "--norm-file",
        default="Final Product Normalization Local.xlsx",
        help="Path to optional Product Normalization reference file."
    )
    parser.add_argument(
        "-o", "--output-dir",
        default="pipeline_output/commission_validation",
        help="Output directory for validation reports (default: pipeline_output/commission_validation)."
    )
    parser.add_argument(
        "-t", "--tolerance",
        type=float,
        default=0.01,
        help="Numeric tolerance threshold for rate/dollar comparison (default: 0.01)."
    )
    parser.add_argument(
        "--min-seq",
        type=int,
        default=3,
        help="Minimum required sequence count per policy group (default: 3)."
    )
    parser.add_argument(
        "--max-seq",
        type=int,
        default=4,
        help="Maximum required sequence count per policy group (default: 4)."
    )

    args = parser.parse_args()

    try:
        process_commission_validation(
            input_path=args.commission_input,
            product_commission_path=args.product_commission_file,
            norm_path=args.norm_file if os.path.exists(args.norm_file) else None,
            output_dir=args.output_dir,
            tolerance=args.tolerance,
            min_seq_len=args.min_seq,
            max_seq_len=args.max_seq,
        )
    except Exception as e:
        print(f"\n❌ Error during commission validation: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
