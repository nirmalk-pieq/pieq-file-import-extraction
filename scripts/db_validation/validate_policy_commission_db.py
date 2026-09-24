#!/usr/bin/env python3
"""
2-Stage Policy Commission Import Validation & DB Verification Engine
====================================================================
Stage 1 (First Validation & Split Normalization):
  - Normalizes product names & updates agent IDs.
  - Performs sequence rate/amount differential roll-ups (normalize_commission_splits.py).
  - Calculates line-item normalized split share percentages (summing to 100.00%).
  - Populates total Carrier Commission and Payout Method (PERCENTAGE / FIXED FEE).
  - Validates total Carrier Commission against master rate sheet (validate_extracted_commission.py).
  - EXPORTS PHASE 1 REPORTS:
      * file_validated_clean_<timestamp>.xlsx
      * file_validated_errors_<timestamp>.xlsx

Stage 2 (Database Verification for agency_id = 5):
  - Takes Phase 1 Valid records and queries PostgreSQL database:
      1. products table (agency_id=5, name={product_name}) -> obtained_product_uuid
      2. policies table (agency_id=5, number={policynumber}, product_id={obtained_product_uuid}) -> obtained_policy_uuid
      3. policy_commission table (agency_id=5, policy_uuid={obtained_policy_uuid}) -> DB commission rules
  - Verifies DB rules against file records.
  - EXPORTS PHASE 2 REPORTS:
      * db_matched_commission_<timestamp>.xlsx
      * db_unmatched_commission_<timestamp>.xlsx

Usage:
  python validate_policy_commission_db.py --file /path/to/commission_upload.csv
  python validate_policy_commission_db.py --dry-run
  python validate_policy_commission_db.py --env-file .env --agency-id 5
"""

import argparse
import datetime
import functools
import math
import os
import re
import sys
import time
from typing import Dict, List, Optional, Set, Tuple, Union

import numpy as np
import pandas as pd
import json

# Import tabulate with fallback
try:
    from tabulate import tabulate
except ImportError:
    def tabulate(data, headers=None, tablefmt=None):
        """Very simple markdown table renderer fallback.
        data: list of rows (iterables)
        headers: list of column names
        """
        if headers:
            lines = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
        else:
            lines = []
        for row in data:
            lines.append("| " + " | ".join(str(v) for v in row) + " |")
        return "\n".join(lines)



# Check psycopg2 / DB driver availability
DB_DRIVER_AVAILABLE = False
try:
    import psycopg2
    from psycopg2 import sql
    from psycopg2.extras import RealDictCursor
    DB_DRIVER_AVAILABLE = True
except ImportError:
    try:
        import psycopg2_binary as psycopg2
        from psycopg2 import sql
        from psycopg2.extras import RealDictCursor
        DB_DRIVER_AVAILABLE = True
    except ImportError:
        DB_DRIVER_AVAILABLE = False





# ---------------------------------------------------------------------------
# Helper: Load .env File
# ---------------------------------------------------------------------------
def load_env_file(env_file_path: str = ".env", override: bool = False) -> int:
    """Loads environment variables from a .env file into os.environ."""
    target_path = env_file_path
    if not os.path.isabs(target_path) and not os.path.exists(target_path):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        alt_path = os.path.join(script_dir, env_file_path)
        if os.path.exists(alt_path):
            target_path = alt_path

    if not os.path.exists(target_path):
        return 0

    count = 0
    try:
        with open(target_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip().strip("'").strip('"')
                    if key:
                        if override or key not in os.environ:
                            os.environ[key] = value
                            count += 1
    except Exception as e:
        print(f"⚠️ Warning: Could not load env file '{target_path}': {e}")
    return count


# ---------------------------------------------------------------------------
# Formatting & Parsing Helpers
# ---------------------------------------------------------------------------
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
    """Parses percentage strings or numeric values into float percentage scale (e.g. 70.0)."""
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
    """Formats float values to standard currency strings."""
    if val < 0:
        return f"-${abs(val):,.2f}"
    return f"${val:,.2f}"


def format_rate(val: float) -> str:
    """Formats float values into percentage strings."""
    return f"{val:.2f} %"


def normalize_product_str(p: str) -> str:
    """Strips punctuation (/ , - *) and collapses whitespace for product matching."""
    if not p or pd.isna(p):
        return ""
    clean = str(p).replace("/", " ").replace("-", " ").replace(",", " ").replace("*", " ")
    return " ".join(clean.lower().split())


def load_file(file_path: str) -> pd.DataFrame:
    """Loads CSV or Excel file (or directory of files) into a unified DataFrame."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File or directory not found: '{file_path}'")

    if os.path.isdir(file_path):
        dfs = []
        for root, _, files in os.walk(file_path):
            for f in sorted(files):
                if f.startswith("~$") or f.startswith("."):
                    continue
                ext = os.path.splitext(f.lower())[1]
                if ext in (".csv", ".xlsx", ".xls"):
                    full_p = os.path.join(root, f)
                    try:
                        dfs.append(load_file(full_p))
                    except Exception as err:
                        print(f"⚠️ Warning: Could not read file '{full_p}': {err}")
        if not dfs:
            raise ValueError(f"No valid CSV or Excel files found in directory '{file_path}'.")
        return pd.concat(dfs, ignore_index=True)

    _, ext = os.path.splitext(file_path.lower())
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(file_path)
    elif ext == ".csv":
        return pd.read_csv(file_path, low_memory=False)
    else:
        raise ValueError(f"Unsupported file format '{ext}' for file '{file_path}'.")


def load_product_normalization(norm_path: str = None) -> Dict[str, str]:
    """Loads product normalization map (Raw Product -> Normalized Product)."""
    if not norm_path or not os.path.exists(norm_path):
        return {}

    try:
        df_norm = load_file(norm_path)
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


def load_agent_roster(roster_path: str = None) -> Dict[str, str]:
    """Loads Agent Roster mapping dictionary (Cleaned Agent Name -> Agent ID)."""
    if not roster_path or not os.path.exists(roster_path):
        return {}

    try:
        df_roster = load_file(roster_path)
        df_roster = df_roster.rename(columns=lambda c: str(c).strip())

        name_col = next((c for c in df_roster.columns if "agent name" in c.lower() or "name" in c.lower()), None)
        id_col = next((c for c in df_roster.columns if "agent id" in c.lower() or "id" in c.lower()), None)

        if not name_col or not id_col:
            return {}

        roster_map = {}
        for _, row in df_roster.iterrows():
            name_val = str(row[name_col]).strip()
            id_val = str(row[id_col]).strip()
            if name_val and id_val and id_val.lower() not in ("nan", "none", ""):
                clean_n = " ".join(name_val.lower().split())
                roster_map[clean_n] = id_val
        return roster_map
    except Exception as e:
        print(f"⚠️ Warning: Could not load agent roster file: {e}")
        return {}


def unpivot_product_commission_rates(excel_path: str) -> pd.DataFrame:
    """Unpivots month tier column groups from ProductWise Commission Rate.xlsx into rules DataFrame."""
    df_raw = load_file(excel_path)
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
                        to_m = 9999
                except Exception:
                    continue

                pct_val = float(row.get(pct_col, 0.0)) if pd.notna(row.get(pct_col)) else 0.0
                dollar_val = float(row.get(dollar_col, 0.0)) if pd.notna(row.get(dollar_col)) else 0.0
                is_fixed = str(row.get(fixed_col, "N")).strip().upper()

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

    return pd.DataFrame(rules)


# ---------------------------------------------------------------------------
# SEQUENCE DIFFERENTIAL ROLL-UP & SPLIT NORMALIZATION (normalize_commission_splits.py)
# ---------------------------------------------------------------------------
def rollup_commission_splits(df: pd.DataFrame) -> pd.DataFrame:
    """
    Performs sequence differential roll-up and percentage split normalization.
    Computes rate differentials (Seq 1=10%, Seq 2=20%-10%=10%, Seq 3=30%-20%=10%) or
    amount differentials (Seq 1=$20, Seq 2=$50-$20=$30, Seq 3=$80-$50=$30),
    normalizes split shares so they sum to 100.00%, and populates total Carrier Commission.
    """
    df_out = df.copy()

    pol_col = next((c for c in df_out.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    prod_col = next((c for c in df_out.columns if str(c).strip().lower() in ("product", "product name", "product_description", "product_name")), None)
    fm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("from month", "from_month", "month")), None)
    seq_col = next((c for c in df_out.columns if str(c).strip().lower() in ("sequence", "seq")), None)
    rate_col = next((c for c in df_out.columns if str(c).strip().lower() in ("rate", "rate %", "commission %")), None)
    amt_col = next((c for c in df_out.columns if str(c).strip().lower() in ("amount", "fee amount", "fee_amount")), None)

    if not pol_col:
        return df_out

    if "Carrier Commission" not in df_out.columns:
        df_out["Carrier Commission"] = "$0.00"
    if "Payout Method" not in df_out.columns:
        df_out["Payout Method"] = ""

    df_out["_amt_num"] = df_out[amt_col].apply(parse_amount) if amt_col else 0.0
    df_out["_rate_num"] = df_out[rate_col].apply(parse_rate) if rate_col else 0.0
    df_out["_seq_num"] = pd.to_numeric(df_out[seq_col], errors="coerce").fillna(1).astype(int) if seq_col else 1

    df_out["_pol_clean"] = clean_key(df_out[pol_col])
    df_out["_prod_clean"] = clean_key(df_out[prod_col]) if prod_col else ""
    df_out["_fm_clean"] = clean_key(df_out[fm_col]) if fm_col else ""

    group_cols = ["_pol_clean"]
    if prod_col:
        group_cols.append("_prod_clean")
    if fm_col:
        group_cols.append("_fm_clean")

    grouped = df_out.groupby(group_cols, sort=False)

    res_amt = list(df_out[amt_col]) if amt_col else []
    res_rate = list(df_out[rate_col]) if rate_col else []
    res_carrier_comm = ["$0.00"] * len(df_out)
    res_payout_method = [""] * len(df_out)
    res_split_share = [0.0] * len(df_out)
    res_diff_val = [""] * len(df_out)

    for _, group in grouped:
        indices = group.index.tolist()
        rates = group["_rate_num"].tolist()
        amounts = group["_amt_num"].tolist()
        seqs = group["_seq_num"].tolist()

        has_nonzero_rate = any(r > 0 for r in rates)

        seq_amt_map = {}
        seq_rate_map = {}
        for s, a, r in zip(seqs, amounts, rates):
            if s not in seq_amt_map:
                seq_amt_map[s] = a
            if s not in seq_rate_map:
                seq_rate_map[s] = r

        max_seq = max(seqs) if seqs else 1

        if not has_nonzero_rate:
            # FIXED FEE ($) — Differential Amounts & 100% Split Shares
            adj_amounts = []
            for s, a in zip(seqs, amounts):
                if a > 0:
                    prev_seq = s - 1
                    if prev_seq in seq_amt_map:
                        prev_amt = seq_amt_map[prev_seq]
                        adj_amounts.append(max(0.0, a - prev_amt))
                    else:
                        adj_amounts.append(a)
                else:
                    adj_amounts.append(a)

            total_amt = seq_amt_map.get(max_seq, max(amounts) if amounts else 0.0)

            rates_calc = []
            if total_amt > 0:
                non_highest_rate_sum = 0.0
                highest_seq_indices = [i for i, s in enumerate(seqs) if s == max_seq]
                primary_highest_idx = highest_seq_indices[-1]

                for i, adj_a in enumerate(adj_amounts):
                    if i != primary_highest_idx:
                        raw_pct = (adj_a / total_amt) * 100.0
                        trunc_pct = math.floor(raw_pct * 100.0) / 100.0
                        rates_calc.append(trunc_pct)
                        non_highest_rate_sum += trunc_pct
                    else:
                        rates_calc.append(None)

                highest_rate = round(100.00 - non_highest_rate_sum, 2)
                rates_calc[primary_highest_idx] = highest_rate
            else:
                rates_calc = [0.0] * len(indices)

            carrier_comm_str = format_amount(total_amt)
            for global_idx, adj_a, r_val in zip(indices, adj_amounts, rates_calc):
                if amt_col:
                    res_amt[global_idx] = format_amount(adj_a)
                if rate_col:
                    res_rate[global_idx] = format_rate(r_val)
                res_diff_val[global_idx] = format_amount(adj_a)
                res_split_share[global_idx] = r_val
                res_carrier_comm[global_idx] = carrier_comm_str
                res_payout_method[global_idx] = "FIXED FEE"

        else:
            # PERCENTAGE (%) — Differential Rates & 100% Split Shares
            rate_diffs = []
            for s, r in zip(seqs, rates):
                prev_seq = s - 1
                if prev_seq in seq_rate_map:
                    prev_r = seq_rate_map[prev_seq]
                    rate_diffs.append(max(0.0, r - prev_r) if r > prev_r else r)
                else:
                    rate_diffs.append(r)

            total_rate = seq_rate_map.get(max_seq, max(rates) if rates else 0.0)

            rates_calc = []
            if total_rate > 0:
                non_highest_rate_sum = 0.0
                highest_seq_indices = [i for i, s in enumerate(seqs) if s == max_seq]
                primary_highest_idx = highest_seq_indices[-1]

                for i, r_diff in enumerate(rate_diffs):
                    if i != primary_highest_idx:
                        raw_pct = (r_diff / total_rate) * 100.0
                        trunc_pct = math.floor(raw_pct * 100.0) / 100.0
                        rates_calc.append(trunc_pct)
                        non_highest_rate_sum += trunc_pct
                    else:
                        rates_calc.append(None)

                highest_rate = round(100.00 - non_highest_rate_sum, 2)
                rates_calc[primary_highest_idx] = highest_rate
            else:
                rates_calc = [0.0] * len(indices)

            carrier_comm_str = format_rate(total_rate)
            for global_idx, r_diff, r_val in zip(indices, rate_diffs, rates_calc):
                if rate_col:
                    res_rate[global_idx] = format_rate(r_val)
                res_diff_val[global_idx] = format_rate(r_diff)
                res_split_share[global_idx] = r_val
                res_carrier_comm[global_idx] = carrier_comm_str
                res_payout_method[global_idx] = "PERCENTAGE"

    if amt_col:
        df_out[amt_col] = res_amt
    if rate_col:
        df_out[rate_col] = res_rate
    df_out["Normalized_Split_Share_%"] = res_split_share
    df_out["Differential_Value"] = res_diff_val
    df_out["Carrier Commission"] = res_carrier_comm
    df_out["Payout Method"] = res_payout_method

    df_out = df_out.drop(columns=["_amt_num", "_rate_num", "_seq_num", "_pol_clean", "_prod_clean", "_fm_clean"], errors="ignore")
    return df_out


# ---------------------------------------------------------------------------
# STAGE 1: File Pre-Validation Engine
# ---------------------------------------------------------------------------
def stage1_pre_validate_dataset(
    df_comm: pd.DataFrame,
    df_rules: Optional[pd.DataFrame],
    norm_dict: Dict[str, str] = None,
    agent_roster_dict: Dict[str, str] = None,
    tolerance: float = 0.01,
) -> Tuple[pd.DataFrame, pd.DataFrame, dict]:
    """
    Stage 1: Applies sequence rollups and validates carrier rates against master rate sheet.
    Returns: (df_stage1_valid, df_stage1_invalid, summary_dict)
    """
    df_norm_splits = rollup_commission_splits(df_comm)
    df_out = df_norm_splits.copy()

    # Enrich / normalize Agent ID whenever Agent Name matches in agent_roster_dict
    agent_name_col = next((c for c in df_out.columns if "agent name" in str(c).strip().lower() or str(c).strip().lower() == "agent_name"), None)
    
    if agent_roster_dict:
        agent_id_col = next((c for c in df_out.columns if str(c).strip().lower() in ("agent id", "agent_id", "agentid")), None)
        if agent_name_col:
            if not agent_id_col:
                df_out["Agent ID"] = ""
                agent_id_col = "Agent ID"
            for idx in df_out.index:
                raw_n = str(df_out.loc[idx, agent_name_col]).strip()
                clean_n = " ".join(raw_n.lower().split())
                if clean_n in agent_roster_dict:
                    df_out.loc[idx, agent_id_col] = agent_roster_dict[clean_n]

    pol_col = next((c for c in df_out.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    prod_col = next((c for c in df_out.columns if str(c).strip().lower() in ("product", "product name", "product_description", "product_name")), None)
    fm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("from month", "from_month", "month", "month from", "month_from")), None)
    tm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("to month", "to_month", "month to", "month_to")), None)
    carrier_comm_col = "Carrier Commission"
    payout_method_col = "Payout Method"

    if not prod_col:
        raise KeyError("Missing Product column in input file.")

    if not fm_col:
        df_out["From_Month"] = 1
        fm_col = "From_Month"
        
    if not tm_col:
        df_out["To_Month"] = 0
        tm_col = "To_Month"

    norm_dict = norm_dict or {}

    validation_statuses = [""] * len(df_out)
    expected_values = [""] * len(df_out)
    variances = [0.0] * len(df_out)
    error_reasons = [""] * len(df_out)
    normalized_product_names = [""] * len(df_out)

    valid_count = 0
    mismatch_val_count = 0
    unmatched_prod_count = 0
    unmatched_month_count = 0

    df_out["_pol_clean"] = clean_key(df_out[pol_col]) if pol_col else pd.Series(["ROW_" + str(i) for i in range(len(df_out))], index=df_out.index)
    df_out["_prod_clean"] = clean_key(df_out[prod_col])
    df_out["_fm_clean"] = clean_key(df_out[fm_col])
    df_out["_tm_clean"] = clean_key(df_out[tm_col])

    grouped = df_out.groupby(["_pol_clean", "_prod_clean", "_fm_clean", "_tm_clean"], sort=False)

    for _, group in grouped:
        indices = group.index.tolist()
        seq_1_idx = indices[0]
        seq_1_row = group.loc[seq_1_idx]

        raw_prod = str(seq_1_row.get(prod_col, "")).strip()
        raw_fm = seq_1_row.get(fm_col, 1)

        clean_prod = " ".join(raw_prod.lower().split())
        norm_prod = normalize_product_str(raw_prod)

        final_norm_prod = raw_prod
        if clean_prod in norm_dict:
            final_norm_prod = norm_dict[clean_prod]
            clean_prod = " ".join(final_norm_prod.lower().split())
            norm_prod = normalize_product_str(final_norm_prod)
        elif norm_prod in norm_dict:
            final_norm_prod = norm_dict[norm_prod]
            clean_prod = " ".join(final_norm_prod.lower().split())
            norm_prod = normalize_product_str(final_norm_prod)

        for idx in indices:
            normalized_product_names[idx] = final_norm_prod

        try:
            from_month_num = int(float(str(raw_fm).strip().replace("'", "").replace('"', "")))
        except (ValueError, TypeError):
            from_month_num = 1

        extracted_comm_str = str(seq_1_row.get(carrier_comm_col, ""))
        payout_method = str(seq_1_row.get(payout_method_col, "")).strip().upper()

        extracted_comm_amt = parse_amount(extracted_comm_str)
        extracted_comm_pct = parse_rate(extracted_comm_str)

        # -------------------------------------------------------------------
        # CHECK 0: MLB Direct Exceptions (Bypass Stage 1 Validation)
        # -------------------------------------------------------------------
        if agent_name_col:
            agent_name_raw = " ".join(str(seq_1_row.get(agent_name_col, "")).lower().split())
            if agent_name_raw in ("main line benefits", "mlb parallel run", "mlb"):
                valid_count += len(indices)
                for idx in indices:
                    validation_statuses[idx] = "VALID"
                    expected_values[idx] = "MLB_EXCEPTION"
                    variances[idx] = 0.0
                    error_reasons[idx] = ""
                continue

        # -------------------------------------------------------------------
        # CHECK 0.5: Negative Split Share Verification
        # -------------------------------------------------------------------
        group_split_shares = [float(group.loc[idx, "Normalized_Split_Share_%"]) for idx in indices]
        has_negative_split = any(s < 0.0 for s in group_split_shares)
        if has_negative_split:
            for idx in indices:
                validation_statuses[idx] = "NEGATIVE_SPLIT_SHARE"
                expected_values[idx] = ">= 0.00 %"
                variances[idx] = min(group_split_shares)
                error_reasons[idx] = f"Group contains negative split share ({min(group_split_shares)}%)."
            continue

        # -------------------------------------------------------------------
        # CHECK 1: Group Split Shares Sum Verification (100.00%)
        # -------------------------------------------------------------------
        group_split_shares = [float(group.loc[idx, "Normalized_Split_Share_%"]) for idx in indices]
        split_sum = round(sum(group_split_shares), 2)
        if abs(split_sum - 100.00) > 0.05:
            group_val = extracted_comm_amt if payout_method == "FIXED FEE" else extracted_comm_pct
            
            if group_val == 0.0:
                for idx in indices:
                    validation_statuses[idx] = "ZERO_COMM_IGNORED"
                    expected_values[idx] = "100.00 %"
                    variances[idx] = round(abs(split_sum - 100.00), 2)
                    error_reasons[idx] = f"Ignored: Commission is 0 and split sum is {split_sum}%"
            else:
                for idx in indices:
                    validation_statuses[idx] = "INVALID_SPLIT_SUM"
                    expected_values[idx] = "100.00 %"
                    variances[idx] = round(abs(split_sum - 100.00), 2)
                    error_reasons[idx] = f"Group split share sum is {split_sum}% instead of expected 100.00%."
            continue

        # -------------------------------------------------------------------
        # CHECK 2: Policy Split Row Count Verification (3 to 4 rows per policy)
        # -------------------------------------------------------------------
        row_count = len(indices)
        if row_count < 3 or row_count > 4:
            for idx in indices:
                validation_statuses[idx] = "INVALID_POLICY_ROW_COUNT"
                expected_values[idx] = "3 to 4 split rows"
                variances[idx] = float(row_count)
                error_reasons[idx] = f"Policy group has {row_count} split rows instead of expected 3 to 4 rows."
            continue

        if df_rules is None or df_rules.empty:
            valid_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "VALID"
                expected_values[idx] = "SKIPPED_RATE_SHEET"
                variances[idx] = 0.0
                error_reasons[idx] = ""
            continue

        p_rules = df_rules[df_rules["Product_Description"] == raw_prod]
        if p_rules.empty:
            p_rules = df_rules[df_rules["Product_Description"] == final_norm_prod]
        if p_rules.empty:
            p_rules = df_rules[df_rules["_prod_clean"] == clean_prod]
        if p_rules.empty and "_prod_norm" in df_rules.columns:
            p_rules = df_rules[df_rules["_prod_norm"] == norm_prod]
        if p_rules.empty:
            p_rules = df_rules[df_rules["_prod_clean"].apply(lambda p: p in clean_prod or clean_prod in p)]

        if p_rules.empty:
            unmatched_prod_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "UNMATCHED_PRODUCT_IN_RATE_SHEET"
                expected_values[idx] = "N/A"
                variances[idx] = 0.0
                error_reasons[idx] = f"Product '{raw_prod}' not found in master rate sheet."
            continue

        m_rules = p_rules[(p_rules["From_Month"] <= from_month_num) & (from_month_num <= p_rules["To_Month"])]
        if m_rules.empty:
            unmatched_month_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "UNMATCHED_MONTH_RANGE"
                expected_values[idx] = "N/A"
                variances[idx] = 0.0
                error_reasons[idx] = f"Month {from_month_num} is out of rate tier ranges for product '{raw_prod}'."
            continue

        agency_rules = m_rules[m_rules["_contract_clean"] == "agency"]
        if agency_rules.empty or (agency_rules["Rate_Dollar"].max() == 0.0 and agency_rules["Rate_Percentage"].max() == 0.0):
            agency_rules = m_rules.sort_values(by=["Rate_Dollar", "Rate_Percentage"], ascending=False)

        agency_rule = agency_rules.sort_values(by=["Rate_Dollar", "Rate_Percentage"], ascending=False).iloc[0]
        exp_agency_pct = agency_rule["Rate_Percentage"]
        exp_agency_dollar = agency_rule["Rate_Dollar"]

        group_valid = True
        group_var = 0.0
        group_exp_str = ""
        group_reason = ""
        group_status = "VALID"

        if payout_method == "PERCENTAGE" or (exp_agency_pct > 0.0 and exp_agency_dollar == 0.0):
            group_exp_str = format_rate(exp_agency_pct)
            group_var = round(abs(extracted_comm_pct - exp_agency_pct), 4)

            if group_var > tolerance:
                group_status = "VALUE_MISMATCH"
                group_reason = f"Carrier Commission mismatch: Extracted benchmark rate '{format_rate(extracted_comm_pct)}' vs expected Agency benchmark '{group_exp_str}'"
                group_valid = False
        elif payout_method == "FIXED FEE" or (exp_agency_dollar > 0.0):
            group_exp_str = format_amount(exp_agency_dollar)
            group_var = round(abs(extracted_comm_amt - exp_agency_dollar), 2)

            if group_var > tolerance:
                group_status = "VALUE_MISMATCH"
                group_reason = f"Carrier Commission mismatch: Extracted benchmark amount '{format_amount(extracted_comm_amt)}' vs expected Agency benchmark '{group_exp_str}'"
                group_valid = False

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
            for idx in indices:
                validation_statuses[idx] = group_status
                expected_values[idx] = group_exp_str
                variances[idx] = group_var
                error_reasons[idx] = group_reason

    df_out["Normalized_Product_Name"] = normalized_product_names
    df_out["Stage1_Status"] = validation_statuses
    df_out["Stage1_Expected_Commission"] = expected_values
    df_out["Stage1_Error_Reason"] = error_reasons

    df_out = df_out.drop(columns=["_pol_clean", "_prod_clean", "_fm_clean"], errors="ignore")

    valid_mask = df_out["Stage1_Status"] == "VALID"
    zero_mask = df_out["Stage1_Status"] == "ZERO_COMM_IGNORED"
    
    invalid_mask = ~(valid_mask | zero_mask)
    
    df_stage1_valid = df_out[valid_mask].copy()
    df_stage1_zero = df_out[zero_mask].copy()
    
    # Grab the original raw rows from df_comm for errors
    df_stage1_invalid = df_comm[invalid_mask].copy()
    
    # Attach just the error context columns back to the raw rows
    df_stage1_invalid["Stage1_Status"] = df_out.loc[invalid_mask, "Stage1_Status"]
    df_stage1_invalid["Stage1_Expected_Commission"] = df_out.loc[invalid_mask, "Stage1_Expected_Commission"]
    df_stage1_invalid["Stage1_Error_Reason"] = df_out.loc[invalid_mask, "Stage1_Error_Reason"]

    summary = {
        "total_input_records": len(df_comm),
        "stage1_valid_records": len(df_stage1_valid),
        "stage1_invalid_records": len(df_stage1_invalid),
        "stage1_zero_ignored": len(df_stage1_zero),
        "value_mismatches": mismatch_val_count,
        "unmatched_products": unmatched_prod_count,
        "unmatched_month_ranges": unmatched_month_count,
    }

    return df_stage1_valid, df_stage1_invalid, df_stage1_zero, summary


# ---------------------------------------------------------------------------
# STAGE 2: Database Verification Engine (PostgreSQL / Read-Only)
# ---------------------------------------------------------------------------
class DatabaseManager:
    """Manages read-only PostgreSQL queries for products, policies, and policy_commission tables."""

    def __init__(self, db_url: str = None, host: str = None, port: int = 5432, dbname: str = None, user: str = None, password: str = None):
        self.db_url = db_url or os.getenv("DATABASE_URL")
        self.user = user or os.getenv("DB_USER")
        self.password = password or os.getenv("DB_PASSWORD")
        self.host = host or os.getenv("DB_HOST", "localhost")
        self.port = port or int(os.getenv("DB_PORT", 5432))
        self.dbname = dbname or os.getenv("DB_NAME", "pieq_db")
        self.conn = None

    def connect(self) -> bool:
        if not DB_DRIVER_AVAILABLE:
            print("⚠️ Notice: psycopg2 driver not installed. Database verification running in DRY-RUN mode.")
            return False

        try:
            target_info = ""
            if self.db_url:
                clean_url = str(self.db_url).strip()
                if clean_url.lower().startswith("jdbc:"):
                    clean_url = clean_url[5:]
                conn_kwargs = {"dsn": clean_url}
                if self.user:
                    conn_kwargs["user"] = self.user
                if self.password:
                    conn_kwargs["password"] = self.password

                host_match = re.search(r"@([^:/]+)", clean_url)
                db_match = re.search(r"/([^/?]+)", clean_url)
                log_host = host_match.group(1) if host_match else "remote host"
                log_db = db_match.group(1) if db_match else "db"
                target_info = f"Host: '{log_host}', DB: '{log_db}', User: '{self.user or 'default'}'"

                self.conn = psycopg2.connect(**conn_kwargs)
            else:
                target_info = f"Host: '{self.host}:{self.port}', DB: '{self.dbname}', User: '{self.user}'"
                self.conn = psycopg2.connect(
                    host=self.host, port=self.port, dbname=self.dbname, user=self.user, password=self.password
                )

            self.conn.set_session(readonly=True, autocommit=True)

            db_version = "PostgreSQL"
            try:
                with self.conn.cursor() as cur:
                    cur.execute("SELECT version();")
                    row = cur.fetchone()
                    if row and row[0]:
                        ver_str = str(row[0])
                        db_version = ver_str.split(",")[0] if "," in ver_str else ver_str.split("\n")[0]
            except Exception:
                pass

            print(f"   🟢 Connected to PostgreSQL ({target_info})")
            print(f"   ℹ️ Server Info: {db_version}")
            return True

        except Exception as e:
            print(f"   ❌ Could not connect to PostgreSQL database: {e}")
            self.conn = None
            return False

    def close(self):
        if self.conn:
            try:
                self.conn.close()
            except Exception:
                pass
            self.conn = None

    @functools.lru_cache(maxsize=4096)
    def get_product_uuid(self, agency_id: int, product_name: str) -> Optional[str]:
        """Queries products table: SELECT * FROM products WHERE agency_id = %s AND (name = %s OR lower(name) = lower(%s))"""
        if not self.conn:
            return None
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                query = """
                    SELECT * FROM products 
                    WHERE agency_id = %s 
                      AND (name = %s OR lower(name) = lower(%s))
                    LIMIT 1;
                """
                cur.execute(query, (agency_id, product_name, product_name))
                row = cur.fetchone()
                if row:
                    val = row.get("uuid") or row.get("id")
                    return str(val) if val is not None else None
        except Exception as e:
            print(f"⚠️ Error querying products for '{product_name}': {e}")
        return None

    @functools.lru_cache(maxsize=16384)
    def get_policy_uuid(self, agency_id: int, policy_number: str, product_uuid: str) -> Optional[str]:
        """Queries policies table: SELECT * FROM policies WHERE agency_id = %s AND number = %s AND product_id::text = %s"""
        if not self.conn:
            return None
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                query = """
                    SELECT * FROM policies 
                    WHERE agency_id = %s 
                      AND number = %s 
                      AND product_id::text = %s
                    LIMIT 1;
                """
                cur.execute(query, (agency_id, policy_number, str(product_uuid)))
                row = cur.fetchone()
                if row:
                    val = row.get("uuid") or row.get("id")
                    return str(val) if val is not None else None
        except Exception as e:
            print(f"⚠️ Error querying policies for number '{policy_number}': {e}")
        return None

    def get_policy_commission(self, agency_id: int, policy_uuid: str) -> List[dict]:
        """Queries policy_commission table: SELECT * FROM policy_commission WHERE agency_id = %s AND policy_uuid::text = %s"""
        if not self.conn:
            return []
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                query = """
                    SELECT * 
                    FROM policy_commission 
                    WHERE agency_id = %s 
                      AND policy_uuid::text = %s;
                """
                cur.execute(query, (agency_id, str(policy_uuid)))
                return [dict(r) for r in cur.fetchall()]
        except Exception as e:
            print(f"⚠️ Error querying policy_commission for policy_uuid '{policy_uuid}': {e}")
        return []

    @functools.lru_cache(maxsize=8192)
    def get_user_uuid(self, agency_id: int, user_code: str = None, user_name: str = None) -> Optional[str]:
        """Queries users and contacts tables to find user UUID (relationship_id for contacts) by user_code (agent_id) or user_name."""
        if not self.conn:
            return None
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                if user_code:
                    cur.execute(
                        "SELECT uuid FROM users WHERE agency_id = %s AND (user_code = %s OR agent_id = %s) LIMIT 1;",
                        (agency_id, str(user_code).strip(), str(user_code).strip())
                    )
                    row = cur.fetchone()
                    if row and row.get("uuid"):
                        return str(row["uuid"])
                if user_name:
                    norm_name = " ".join(str(user_name).strip().lower().split())
                    cur.execute(
                        "SELECT uuid FROM users WHERE agency_id = %s AND (lower(name) = %s OR lower(full_name) = %s OR lower(user_name) = %s) LIMIT 1;",
                        (agency_id, norm_name, norm_name, norm_name)
                    )
                    row = cur.fetchone()
                    if row and row.get("uuid"):
                        return str(row["uuid"])
                    
                    # Fallback to contacts table: concat first_name and last_name, take relationship_id as user_uuid
                    cur.execute(
                        "SELECT relationship_id FROM contacts WHERE agency_id = %s AND (lower(concat(first_name, ' ', last_name)) = %s OR lower(first_name) = %s OR lower(last_name) = %s) AND relationship_id IS NOT NULL LIMIT 1;",
                        (agency_id, norm_name, norm_name, norm_name)
                    )
                    row = cur.fetchone()
                    if row and row.get("relationship_id"):
                        return str(row["relationship_id"])
        except Exception as e:
            pass
        return None

    def fetch_all_users_map(self, agency_id: int) -> Dict[str, str]:
        """Fetches a map of (clean_user_code or clean_user_name) -> user_uuid for the agency from users and contacts tables."""
        if not self.conn:
            return {}
        user_map = {}
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    SELECT column_name 
                    FROM information_schema.columns 
                    WHERE table_name = 'users';
                """)
                cols = set(r["column_name"].lower() for r in cur.fetchall())
                
                code_col = "user_code" if "user_code" in cols else ("agent_id" if "agent_id" in cols else None)
                name_col = "name" if "name" in cols else ("full_name" if "full_name" in cols else ("user_name" if "user_name" in cols else None))
                
                select_cols = ["uuid"]
                if code_col:
                    select_cols.append(code_col)
                if name_col:
                    select_cols.append(name_col)
                    
                query = f"SELECT {', '.join(select_cols)} FROM users WHERE agency_id = %s;"
                cur.execute(query, (agency_id,))
                rows = cur.fetchall()
                for r in rows:
                    u_uuid = str(r.get("uuid", ""))
                    if not u_uuid:
                        continue
                    if code_col and r.get(code_col):
                        c_val = str(r[code_col]).strip().lower()
                        if c_val:
                            user_map[c_val] = u_uuid
                    if name_col and r.get(name_col):
                        n_val = " ".join(str(r[name_col]).strip().lower().split())
                        if n_val:
                            user_map[n_val] = u_uuid
                            
                # Query contacts table for fallback mappings (concat first_name + last_name -> relationship_id)
                try:
                    cur.execute("""
                        SELECT relationship_id, first_name, last_name 
                        FROM contacts 
                        WHERE agency_id = %s AND relationship_id IS NOT NULL;
                    """, (agency_id,))
                    c_rows = cur.fetchall()
                    for r in c_rows:
                        rel_id = str(r.get("relationship_id", "")).strip()
                        if not rel_id:
                            continue
                        fname = str(r.get("first_name", "") or "").strip()
                        lname = str(r.get("last_name", "") or "").strip()
                        full = " ".join(f"{fname} {lname}".lower().split())
                        if full and full not in user_map:
                            user_map[full] = rel_id
                except Exception as ce:
                    pass
        except Exception as e:
            print(f"⚠️ Warning: Could not fetch user map from database: {e}")
        return user_map

    def fetch_all_agency_commission_rules(self, agency_id: int) -> Dict[Union[str, Tuple[str, str]], List[dict]]:
        """
        Executes a SINGLE SQL query to pre-fetch ALL products, policies, and policy_commission rules
        for the entire agency into an in-memory lookup map for zero-latency Stage 2 verification.
        """
        if not self.conn:
            return {}
        result_map = {}
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                query = """
                    SELECT 
                        p.uuid AS product_uuid,
                        p.name AS product_name,
                        pol.uuid AS policy_uuid,
                        pol.number AS policy_number,
                        pol.writing_agent_user_id,
                        pol.pay_to_agent_user_id,
                        pc.uuid AS commission_uuid,
                        pc.rule_name,
                        pc.commission_type,
                        pc.payment_frequency,
                        pc.month_from,
                        pc.month_to,
                        pc.payout_method,
                        pc.fee_amount,
                        pc.premium_amount,
                        pc.commission_split
                    FROM products p
                    JOIN policies pol ON (pol.agency_id = p.agency_id AND pol.product_id::text = p.uuid::text)
                    LEFT JOIN policy_commission pc ON (pc.agency_id = p.agency_id AND pc.policy_uuid::text = pol.uuid::text)
                    WHERE p.agency_id = %s
                      AND pol.agency_id = %s;
                """
                cur.execute(query, (agency_id, agency_id))
                rows = cur.fetchall()

                for r in rows:
                    p_num = str(r["policy_number"]).strip()
                    p_name = " ".join(str(r["product_name"]).lower().split())
                    rec_dict = dict(r)

                    # 1. Composite key index for exact product-policy pair matching
                    comp_key = (p_num, p_name)
                    if comp_key not in result_map:
                        result_map[comp_key] = []
                    result_map[comp_key].append(rec_dict)

                    # 2. Simple policy number index for fallback matching
                    if p_num not in result_map:
                        result_map[p_num] = []
                    result_map[p_num].append(rec_dict)

            return result_map
        except Exception as e:
            print(f"⚠️ Warning: Could not bulk fetch agency commission rules: {e}")
            return {}


    def batch_fetch_policy_commission_rules(
        self, agency_id: int, exact_pols: List[str], similar_prefixes: List[str]
    ) -> Dict[Union[str, Tuple[str, str]], List[dict]]:
        """
        Executes a single Partition-First SQL Batch JOIN Query for a chunk of policy numbers.
        Returns dict mapping (policy_number, clean_product_name) -> List of joined DB record dicts.
        """
        if not self.conn or (not exact_pols and not similar_prefixes):
            return {}

        clean_exact = [str(p).strip() for p in exact_pols if str(p).strip()]
        clean_similar = [str(p).strip() for p in similar_prefixes if str(p).strip()]
        
        if not clean_exact and not clean_similar:
            return {}

        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                query_conditions = []
                params = [agency_id, agency_id]
                
                if clean_exact:
                    placeholders = ", ".join(["%s"] * len(clean_exact))
                    query_conditions.append(f"pol.number IN ({placeholders})")
                    params.extend(clean_exact)
                    
                if clean_similar:
                    like_clauses = " OR ".join(["pol.number LIKE %s"] * len(clean_similar))
                    query_conditions.append(f"({like_clauses})")
                    params.extend([f"{p}%" for p in clean_similar])
                    
                where_clause = " OR ".join(query_conditions)
                
                query = f"""
                    SELECT 
                        p.uuid AS product_uuid,
                        p.name AS product_name,
                        pol.uuid AS policy_uuid,
                        pol.number AS policy_number,
                        pol.writing_agent_user_id,
                        pol.pay_to_agent_user_id,
                        pc.uuid AS commission_uuid,
                        pc.rule_name,
                        pc.commission_type,
                        pc.payment_frequency,
                        pc.month_from,
                        pc.month_to,
                        pc.payout_method,
                        pc.fee_amount,
                        pc.premium_amount,
                        pc.commission_split
                    FROM products p
                    JOIN policies pol ON (pol.agency_id = p.agency_id AND pol.product_id::text = p.uuid::text)
                    LEFT JOIN policy_commission pc ON (pc.agency_id = p.agency_id AND pc.policy_uuid::text = pol.uuid::text)
                    WHERE p.agency_id = %s
                      AND pol.agency_id = %s
                      AND ({where_clause});
                """
                cur.execute(query, params)
                rows = cur.fetchall()

                result_map = {}
                for r in rows:
                    p_num = str(r["policy_number"]).strip()
                    p_name = " ".join(str(r["product_name"]).lower().split())
                    rec_dict = dict(r)

                    # 1. Composite key index for exact product-policy pair matching
                    comp_key = (p_num, p_name)
                    if comp_key not in result_map:
                        result_map[comp_key] = []
                    result_map[comp_key].append(rec_dict)

                    # 2. Simple policy number index for fallback matching
                    if p_num not in result_map:
                        result_map[p_num] = []
                    result_map[p_num].append(rec_dict)

                return result_map

        except Exception as e:
            print(f"⚠️ Error executing batch policy JOIN query: {e}")
            return {}


def normalize_level(level_val: str) -> str:
    """
    Normalizes level strings to standard format:
    SUBAGT1 -> SA1, SUBAGT2 -> SA2, SUBAGT3 -> SA3 (and SUBAGT<N> -> SA<N>)
    LVL1 -> Level 1, LVL2 -> Level 2 (and LVL<N> -> Level <N>)
    """
    if not level_val:
        return ""
    val_upper = str(level_val).strip().upper()
    
    # SUBAGT<N> -> SA<N>
    subagt_match = re.match(r"^SUBAGT(\d+)$", val_upper)
    if subagt_match:
        return f"SA{subagt_match.group(1)}"
        
    # LVL<N> -> Level <N>
    lvl_match = re.match(r"^LVL(\d+)$", val_upper)
    if lvl_match:
        return f"Level {lvl_match.group(1)}"
        
    return str(level_val).strip()


def build_file_commission_split(group: pd.DataFrame, seq_col, agent_col, agent_name_col, contract_col, split_col) -> list:
    """
    Builds a commission_split JSON array from file rows (same group = same policy+product),
    in the exact same format as the import script's payload builder.
    Reads role/level from the 'Contract' column directly.
    """
    raw_items = []
    seen_keys = set()
    for idx in group.index:
        row = group.loc[idx]
        seq = int(row.get(seq_col, 0)) if seq_col and pd.notna(row.get(seq_col)) else 0
        agent_id = str(row.get(agent_col, "")).strip() if agent_col else ""
        agent_name = str(row.get(agent_name_col, "")).strip() if agent_name_col else agent_id
        contract = str(row.get(contract_col, "")).strip() if contract_col else ""
        commission_value = float(row.get(split_col, 0.0)) if split_col else 0.0

        # Deduplicate by (sequence, agent_code) — file has multiple From Month rows per agent
        dedup_key = (seq, agent_id)
        if dedup_key in seen_keys:
            continue
        seen_keys.add(dedup_key)

        raw_items.append({
            "sequence": seq,
            "user_name": agent_name,
            "user_code": agent_id if agent_id else None,
            "commission_value": commission_value,
            "contract": contract,
        })

    # Sort descending by sequence (highest seq = order 1 = Agency), matching import script logic
    raw_items.sort(key=lambda x: x["sequence"], reverse=True)

    commission_split = []
    for order_idx, item in enumerate(raw_items):
        order = order_idx + 1
        contract_val = item["contract"]
        norm_level = normalize_level(contract_val)
        
        # Determine role and entity from Contract column and order
        contract_lower = contract_val.lower().strip()
        if order == 1 or contract_lower in ("agency", ""):
            role = "Agency"
            entity = "agency"
        elif "sales" in contract_lower or "leader" in contract_lower or order == 2:
            role = "Sales Leader"
            entity = "agent"
        elif "sub" in contract_lower or order >= 4:
            role = "Sub Agent"
            entity = "agent"
        else:
            role = "Agent"
            entity = "agent"

        if order == 1 or contract_lower in ("agency", ""):
            commission_split.append({
                "order": 1,
                "role": "Agency",
                "user_name": item["user_name"],
                "user_code": "",
                "commission_value": round(item["commission_value"], 2),
                "entity": "agency",
                "level": "",
                "user_uuid": "62a97a01-2801-4c7c-9cd0-cf38f2e0a4f4"
            })
        else:
            commission_split.append({
                "order": order,
                "role": role,
                "user_name": item["user_name"],
                "user_code": item["user_code"],
                "commission_value": round(item["commission_value"], 2),
                "entity": entity,
                "level": norm_level,
            })

    return commission_split

def stage2_db_verify_dataset(
    df_valid: pd.DataFrame,
    db_mgr: DatabaseManager,
    agency_id: int = 5,
    dry_run: bool = False,
    chunk_size: int = 2000,
    tolerance: float = 0.02,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, list]:
    """
    Stage 2: Database Verification for Stage-1 Valid Records.
    Uses chunked Partition-First SQL Batch JOIN Queries (2000 policy numbers per chunk).
    Returns: (matched_df, ambiguous_df, unmatched_df, missing_policy_payloads)
    """
    if df_valid.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), []

    df_out = df_valid.copy().reset_index(drop=True)

    pol_col = next((c for c in df_out.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    prod_col = next((c for c in df_out.columns if str(c).strip().lower() in ("normalized_product_name", "product name", "product")), None)
    rate_col = next((c for c in df_out.columns if str(c).strip().lower() in ("rate", "rate %", "commission %")), None)
    agent_col = next((c for c in df_out.columns if str(c).strip().lower() in ("agent id", "agent_id", "agentid")), None)
    agent_name_col = next((c for c in df_out.columns if "agent name" in str(c).strip().lower() or str(c).strip().lower() == "agent_name"), None)
    seq_col = next((c for c in df_out.columns if str(c).strip().lower() in ("sequence", "seq")), None)
    contract_col = next((c for c in df_out.columns if str(c).strip().lower() in ("contract", "level", "role")), None)
    split_col = "Normalized_Split_Share_%" if "Normalized_Split_Share_%" in df_out.columns else None
    fm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("from month", "from_month", "month", "month from", "month_from")), None)
    tm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("to month", "to_month", "month to", "month_to")), None)

    obtained_product_uuids = [""] * len(df_out)
    obtained_policy_uuids = [""] * len(df_out)
    stage2_statuses = [""] * len(df_out)
    unmatched_reasons = [""] * len(df_out)
    db_vals = [""] * len(df_out)
    file_vals = [""] * len(df_out)

    matched_count = 0
    unmatched_count = 0
    missing_policy_payloads = []
    mismatch_policy_payloads = []
    invalid_user_uuid_payloads = []

    grouped = df_out.groupby([pol_col, prod_col], sort=False) if pol_col and prod_col else df_out.groupby(df_out.index)
    group_items = list(grouped)
    total_groups = len(group_items)

    print(f"🔍 Starting Stage 2 Chunked DB Verification for {total_groups:,} policy groups ({len(df_out):,} rows)...")

    # Extract unique policy numbers for batch SQL chunking
    unique_exact_policy_numbers = []
    unique_similar_prefixes = []
    
    for _, group in group_items:
        if pol_col:
            seq_1_row = group.loc[group.index[0]]
            pol = str(seq_1_row[pol_col]).strip()
            is_mapped = bool(seq_1_row.get("_is_mapped", False)) if "_is_mapped" in group.columns else False
            
            if not is_mapped and len(pol) > 11:
                prefix = pol[:11]
                if prefix not in unique_similar_prefixes:
                    unique_similar_prefixes.append(prefix)
            else:
                if pol not in unique_exact_policy_numbers:
                    unique_exact_policy_numbers.append(pol)

    # Pre-fetch all user UUIDs mapping for agency
    db_user_map = db_mgr.fetch_all_users_map(agency_id) if not dry_run and db_mgr.conn else {}

    # Pre-fetch ALL products, policies, and policy_commission rules in 1 single bulk SQL query
    chunk_db_map = {}
    prefix_db_map = {}
    if not dry_run and db_mgr.conn:
        print(f" ⚡ Pre-fetching ALL agency policies & commission rules in 1 single bulk query for zero-latency verification...")
        chunk_db_map = db_mgr.fetch_all_agency_commission_rules(agency_id)
        print(f"    ✅ Bulk loaded {len(chunk_db_map):,} policy/product indexes into RAM.")
        
        # Pre-index 11-char policy prefixes to eliminate O(N) linear set scans inside verification loop
        for k in chunk_db_map.keys():
            if isinstance(k, str) and len(k) >= 11:
                pref = k[:11]
                if pref not in prefix_db_map:
                    prefix_db_map[pref] = set()
                prefix_db_map[pref].add(k)

    not_found_products = set()
    not_found_policies = set()
    not_found_commissions = set()

    for group_idx, (_, group) in enumerate(group_items, start=1):
        indices = group.index.tolist()
        seq_1_row = group.loc[indices[0]]

        pol_number = str(seq_1_row.get(pol_col, "")).strip()
        prod_name = str(seq_1_row.get(prod_col, "")).strip()
        clean_prod_name = " ".join(prod_name.lower().split())

        if dry_run or not db_mgr.conn:
            mock_prod_uuid = f"mock-prod-{abs(hash(prod_name)) % 1000000:06d}"
            mock_pol_uuid = f"mock-pol-{abs(hash(pol_number)) % 1000000:06d}"

            for idx in indices:
                obtained_product_uuids[idx] = mock_prod_uuid
                obtained_policy_uuids[idx] = mock_pol_uuid
                stage2_statuses[idx] = "MATCHED"
                unmatched_reasons[idx] = ""
                db_vals[idx] = "DRY_RUN_MOCK"
                file_vals[idx] = "DRY_RUN_MOCK"
                matched_count += 1
        else:
            is_mapped = bool(seq_1_row.get("_is_mapped", False)) if "_is_mapped" in seq_1_row else False
            is_ambiguous = False
            
            if not is_mapped and len(pol_number) > 11:
                prefix = pol_number[:11]
                matched_full_pols = prefix_db_map.get(prefix, set())
                
                if len(matched_full_pols) > 1:
                    is_ambiguous = True
                    db_records = []
                elif len(matched_full_pols) == 1:
                    found_pol = next(iter(matched_full_pols))
                    db_records = chunk_db_map.get((found_pol, clean_prod_name), [])
                    if not db_records:
                        db_records = chunk_db_map.get(found_pol, [])
                else:
                    db_records = []
            else:
                # 1. Primary lookup using composite tuple key
                db_records = chunk_db_map.get((pol_number, clean_prod_name), [])
                # 2. Fallback lookup
                if not db_records:
                    pol_records = chunk_db_map.get(pol_number, [])
                    db_records = [
                        r for r in pol_records 
                        if clean_prod_name in " ".join(str(r.get("product_name", "")).lower().split())
                        or " ".join(str(r.get("product_name", "")).lower().split()) in clean_prod_name
                    ]
                    if not db_records and pol_records:
                        db_records = pol_records

            # Fallback DB Queries if not in chunk map
            found_prod_uuid = db_records[0].get("product_uuid") if db_records else None
            found_pol_uuid = db_records[0].get("policy_uuid") if db_records else None

            if not found_prod_uuid and prod_name not in not_found_products:
                found_prod_uuid = db_mgr.get_product_uuid(agency_id, prod_name)
                if not found_prod_uuid:
                    not_found_products.add(prod_name)
            if not found_pol_uuid and found_prod_uuid and pol_number not in not_found_policies:
                found_pol_uuid = db_mgr.get_policy_uuid(agency_id, pol_number, found_prod_uuid)
                if not found_pol_uuid:
                    not_found_policies.add(pol_number)

            if is_ambiguous:
                ambiguous_count += len(indices)
                for idx in indices:
                    obtained_product_uuids[idx] = None
                    obtained_policy_uuids[idx] = None
                    stage2_statuses[idx] = "AMBIGUOUS_SIMILAR_MATCH"
                    unmatched_reasons[idx] = f"Policy number '{pol_number[:11]}...' matched multiple policies in DB."
                    db_vals[idx] = "N/A"
                    file_vals[idx] = "N/A"
            elif not found_prod_uuid or not found_pol_uuid:
                unmatched_count += len(indices)
                for idx in indices:
                    obtained_product_uuids[idx] = found_prod_uuid
                    obtained_policy_uuids[idx] = found_pol_uuid
                    stage2_statuses[idx] = "NO_POLICY_FOUND_IN_DB"
                    unmatched_reasons[idx] = f"Policy number '{pol_number}' not found in policies table."
                    db_vals[idx] = "N/A"
                    file_vals[idx] = "N/A"
            else:
                # Obtain all DB commission rules for policy
                comm_records = [r for r in db_records if r.get("commission_uuid")]
                if not comm_records and found_pol_uuid and found_pol_uuid not in not_found_commissions:
                    comm_records = db_mgr.get_policy_commission(agency_id, found_pol_uuid)
                    if not comm_records:
                        not_found_commissions.add(found_pol_uuid)

                if not comm_records:
                    unmatched_count += len(indices)
                    for idx in indices:
                        row = group.loc[idx]
                        obtained_product_uuids[idx] = found_prod_uuid
                        obtained_policy_uuids[idx] = found_pol_uuid
                        stage2_statuses[idx] = "NO_COMMISSION_RULE_FOUND_IN_DB"
                        unmatched_reasons[idx] = f"No commission rules in policy_commission table for policy_uuid '{found_pol_uuid}'."
                        db_vals[idx] = "N/A"
                        file_vals[idx] = ""
                        
                        raw_fm = row.get(fm_col, 1) if fm_col else 1
                        raw_tm = row.get(tm_col, 0) if tm_col else 0
                        try:
                            file_fm = int(float(str(raw_fm).strip().replace("'", "").replace('"', "")))
                        except:
                            file_fm = 1
                        try:
                            file_tm = int(float(str(raw_tm).strip().replace("'", "").replace('"', "")))
                        except:
                            file_tm = 0
                            
                        expected_db_tm = None if file_tm == 0 else file_tm
                        
                        if fm_col and tm_col:
                            tier_group = group[(group[fm_col] == raw_fm) & (group[tm_col] == raw_tm)]
                        elif fm_col:
                            tier_group = group[group[fm_col] == raw_fm]
                        elif tm_col:
                            tier_group = group[group[tm_col] == raw_tm]
                        else:
                            tier_group = group
                            
                        file_commission_split = build_file_commission_split(
                            tier_group, seq_col, agent_col, agent_name_col, contract_col, split_col
                        )
                        carrier_comm_raw = str(row.get("Carrier Commission", "") or row.get(rate_col, "") or row.get(amt_col, "")).strip()
                        file_payout_method_str = str(row.get("Payout Method", "")).strip().upper()
                        
                        fee_val = 0.0
                        prem_val = 0.0
                        if "FIXED FEE" in file_payout_method_str:
                            fee_val = parse_amount(carrier_comm_raw)
                            payout_method_code = 1
                        else:
                            # Default to PERCENTAGE payout method (code 2)
                            fee_val = parse_rate(carrier_comm_raw)
                            payout_method_code = 2

                        # Enrich file_commission_split items with user_uuid and sanitize values
                        sanitized_splits = []
                        missing_uuid_in_group = False
                        for split_item in file_commission_split:
                            s_item = split_item.copy()
                            c_val = float(s_item.get("commission_value") or 0.0)
                            s_item["commission_value"] = max(0.0, c_val) # Ensure no negative commission values
                            
                            if not s_item.get("user_uuid"):
                                u_code = str(s_item.get("user_code") or "").strip().lower()
                                u_name = " ".join(str(s_item.get("user_name") or "").strip().lower().split())
                                fetched_uuid = db_user_map.get(u_code) if u_code else None
                                if not fetched_uuid and u_name:
                                    fetched_uuid = db_user_map.get(u_name)
                                if not fetched_uuid and (u_code or u_name):
                                    fetched_uuid = db_mgr.get_user_uuid(agency_id, user_code=s_item.get("user_code"), user_name=s_item.get("user_name"))
                                    if u_code:
                                        db_user_map[u_code] = fetched_uuid
                                    if u_name:
                                        db_user_map[u_name] = fetched_uuid
                                if fetched_uuid:
                                    s_item["user_uuid"] = fetched_uuid
                            
                            # Check: if commission_value > 0, user_uuid MUST be present
                            if s_item["commission_value"] > 0 and not s_item.get("user_uuid"):
                                missing_uuid_in_group = True
                            
                            sanitized_splits.append(s_item)
                        
                        missing_payload = {
                            "policyUuid": found_pol_uuid,
                            "policyNumber": pol_number,
                            "ruleName": f"Carrier - Product - COMMISSION",
                            "version": 1,
                            "commissionType": 1,
                            "paymentFrequency": 1,
                            "payoutMethod": payout_method_code,
                            "monthFrom": file_fm,
                            "monthTo": expected_db_tm,
                            "feeAmount": fee_val,
                            "premiumAmount": prem_val,
                            "commissionSplit": sanitized_splits
                        }
                        
                        if missing_uuid_in_group:
                            # Do NOT include in POST payloads list if user_uuid is missing for active commission splits
                            missing_payload["_issue_reason"] = "Contains split with commission_value > 0 but missing user_uuid"
                            invalid_user_uuid_payloads.append(missing_payload)
                        else:
                            missing_policy_payloads.append(missing_payload)
                else:
                    for idx in indices:
                        row = group.loc[idx]
                        obtained_product_uuids[idx] = found_prod_uuid
                        obtained_policy_uuids[idx] = found_pol_uuid

                        # Extract File Months
                        raw_fm = row.get(fm_col, 1) if fm_col else 1
                        raw_tm = row.get(tm_col, 0) if tm_col else 0
                        try:
                            file_fm = int(float(str(raw_fm).strip().replace("'", "").replace('"', "")))
                        except:
                            file_fm = 1
                        try:
                            file_tm = int(float(str(raw_tm).strip().replace("'", "").replace('"', "")))
                        except:
                            file_tm = 0
                            
                        # Normalize File To Month = 0 to DB expected NULL
                        expected_db_tm = None if file_tm == 0 else file_tm

                        # Build file commission_split JSON for this specific month tier
                        if fm_col and tm_col:
                            tier_group = group[(group[fm_col] == raw_fm) & (group[tm_col] == raw_tm)]
                        elif fm_col:
                            tier_group = group[group[fm_col] == raw_fm]
                        elif tm_col:
                            tier_group = group[group[tm_col] == raw_tm]
                        else:
                            tier_group = group
                            
                        file_commission_split = build_file_commission_split(
                            tier_group, seq_col, agent_col, agent_name_col, contract_col, split_col
                        )
                        file_split_json = json.dumps(file_commission_split)

                        row_rate = parse_rate(row.get(rate_col, 0.0)) if rate_col else 0.0
                        row_split = float(row.get("Normalized_Split_Share_%", 0.0))
                        row_agent = str(row.get(agent_col, "")).strip() if agent_col else ""
                        row_agent_name = str(row.get(agent_name_col, "")).strip() if agent_name_col else row_agent
                        norm_row_name = " ".join(row_agent_name.lower().split())

                        match_found = False
                        unmatched_agent = False
                        matched_db_value = None
                        db_split_raw = None
                        matched_commission_uuid = None
                        matched_db_rec = None

                        for db_rec in comm_records:
                            db_fm = db_rec.get("month_from")
                            db_tm = db_rec.get("month_to")
                            
                            # Skip DB rules that do not match this file row's period
                            if db_fm != file_fm or db_tm != expected_db_tm:
                                continue
                                
                            db_split_raw = db_rec.get("commission_split")
                            db_fee = float(db_rec.get("fee_amount") or 0.0)

                            if isinstance(db_split_raw, list):
                                for item in db_split_raw:
                                    if not isinstance(item, dict):
                                        continue
                                        
                                    db_code = str(item.get("user_code") or "").strip()
                                    db_name = str(item.get("user_name") or "").strip()
                                    norm_db_name = " ".join(db_name.lower().split())
                                    
                                    # Match Agent by user_code, fallback to normalized user_name
                                    is_agent_match = False
                                    if row_agent and db_code and row_agent == db_code:
                                        is_agent_match = True
                                    elif norm_row_name and norm_db_name and norm_row_name == norm_db_name:
                                        is_agent_match = True
                                    elif str(item.get("order", "")) == "1" and (
                                        ("main line benefits" in norm_row_name and "mlb parallel run" in norm_db_name) or
                                        ("mlb parallel run" in norm_row_name and "main line benefits" in norm_db_name)
                                    ):
                                        is_agent_match = True
                                        
                                    if is_agent_match:
                                        db_val = item.get("commission_value")
                                        if db_val is not None:
                                            matched_db_value = float(db_val)
                                            if abs(matched_db_value - row_split) <= tolerance or abs(matched_db_value - row_rate) <= tolerance:
                                                match_found = True
                                                matched_commission_uuid = db_rec.get("commission_uuid")
                                                matched_db_rec = db_rec
                                                break
                                        unmatched_agent = True
                                        
                                if match_found:
                                    break
                                else:
                                    # Record the uuid even if it's a mismatch so we know which one to update
                                    matched_commission_uuid = db_rec.get("commission_uuid")
                                    matched_db_rec = db_rec
                                    
                            # Fallback for Fixed Fee rules where db_split_raw might be empty but fee_amount matches
                            if not match_found and not unmatched_agent and db_fee > 0 and abs(parse_amount(row.get(rate_col, 0)) - db_fee) <= tolerance:
                                match_found = True
                                matched_commission_uuid = db_rec.get("commission_uuid")
                                matched_db_rec = db_rec
                                break
                            
                            # If we get here, it means we found a rule for the period but failed the agent/rate check
                            if not match_found:
                                matched_commission_uuid = db_rec.get("commission_uuid")
                                matched_db_rec = db_rec

                        if match_found:
                            stage2_statuses[idx] = "MATCHED"
                            unmatched_reasons[idx] = ""
                            db_vals[idx] = json.dumps(db_split_raw) if isinstance(db_split_raw, (list, dict)) else str(db_split_raw)
                            file_vals[idx] = file_split_json
                            matched_count += 1
                        else:
                            stage2_statuses[idx] = "DB_COMMISSION_VALUE_MISMATCH"
                            if unmatched_agent:
                                matched_db_value_str = f"{matched_db_value}%" if matched_db_value is not None else "None"
                                unmatched_reasons[idx] = (
                                    f"Agent '{row_agent_name}' found in DB, but DB commission_value is {matched_db_value_str} (File has {row_split}%). "
                                )
                            else:
                                code_str = f" (Code: {row_agent})" if row_agent else ""
                                unmatched_reasons[idx] = (
                                    f"No commission rule record found in policy_commission table for agent '{row_agent_name}'{code_str}. "
                                )
                            db_vals[idx] = json.dumps(db_split_raw) if isinstance(db_split_raw, (list, dict)) else str(db_split_raw)
                            file_vals[idx] = file_split_json
                            unmatched_count += 1
                            
                            if matched_commission_uuid and matched_db_rec:
                                # Merge DB fields into file splits and build diff
                                merged_splits = []
                                diff_lines = []
                                
                                db_splits_list = db_split_raw if isinstance(db_split_raw, list) else []
                                db_split_dict = {str(s.get("order", "")): s for s in db_splits_list}
                                
                                for f_split in file_commission_split:
                                    m_split = f_split.copy()
                                    order_str = str(m_split.get("order", ""))
                                    db_s = db_split_dict.get(order_str)
                                    
                                    if db_s:
                                        # Preserve existing DB fields
                                        for key in ["user_uuid", "role", "level", "user_name", "user_code"]:
                                            if key in db_s:
                                                m_split[key] = db_s[key]
                                        
                                        # Diff check (use tolerance parameter)
                                        old_val = float(db_s.get("commission_value") or 0.0)
                                        new_val = float(m_split.get("commission_value") or 0.0)
                                        
                                        if abs(old_val - new_val) > tolerance:
                                            diff_lines.append(f"   Order {order_str} | User: {m_split.get('user_name', 'N/A')} | Rate: {old_val}% ➔ {new_val}%")
                                    else:
                                        diff_lines.append(f"   Order {order_str} | User: [NEW] ➔ {m_split.get('user_name', 'N/A')} | Rate: [NEW] ➔ {m_split.get('commission_value', 0.0)}%")
                                        if "level" in m_split:
                                            m_split["level"] = normalize_level(m_split["level"])
                                        
                                    if str(m_split.get("order", "")) == "1" or str(m_split.get("role", "")).lower() == "agency":
                                        m_split["user_uuid"] = "62a97a01-2801-4c7c-9cd0-cf38f2e0a4f4"
                                        m_split["user_code"] = ""
                                        m_split["level"] = ""

                                    # Lookup user_uuid if missing (e.g. for new orders)
                                    if not m_split.get("user_uuid"):
                                        u_code = str(m_split.get("user_code") or "").strip().lower()
                                        u_name = " ".join(str(m_split.get("user_name") or "").strip().lower().split())
                                        fetched_uuid = db_user_map.get(u_code) if u_code else None
                                        if not fetched_uuid and u_name:
                                            fetched_uuid = db_user_map.get(u_name)
                                        if not fetched_uuid and (u_code or u_name):
                                            fetched_uuid = db_mgr.get_user_uuid(agency_id, user_code=m_split.get("user_code"), user_name=m_split.get("user_name"))
                                            # Cache result in db_user_map to avoid redundant DB queries
                                            if u_code:
                                                db_user_map[u_code] = fetched_uuid
                                            if u_name:
                                                db_user_map[u_name] = fetched_uuid
                                        if fetched_uuid:
                                            m_split["user_uuid"] = fetched_uuid
                                    # Ensure no negative commission values
                                    c_val = float(m_split.get("commission_value") or 0.0)
                                    m_split["commission_value"] = max(0.0, c_val)

                                    merged_splits.append(m_split)
                                    
                                db_fm_val = int(matched_db_rec.get("month_from") or 1)
                                db_tm_val = int(matched_db_rec.get("month_to") or 12) if matched_db_rec.get("month_to") is not None else "None"
                                file_tm_val = expected_db_tm if expected_db_tm is not None else "None"
                                if db_fm_val != file_fm or matched_db_rec.get("month_to") != expected_db_tm:
                                    diff_lines.append(f"   Period: {db_fm_val}-{db_tm_val} ➔ {file_fm}-{file_tm_val}")

                                if not diff_lines:
                                    # If only names were fuzzy matched or rates differed by <0.00%, don't bother PUTting
                                    continue
                                
                                diff_text = "\n".join(diff_lines)

                                mismatch_payload = {
                                    "rule_uuid": matched_commission_uuid,
                                    "policyNumber": pol_number,
                                    "policyUuid": matched_db_rec.get("policy_uuid"),
                                    "ruleName": matched_db_rec.get("rule_name", f"Carrier - Product - COMMISSION"),
                                    "version": matched_db_rec.get("version", 1),
                                    "commissionType": matched_db_rec.get("commission_type", 1),
                                    "paymentFrequency": matched_db_rec.get("payment_frequency", 1),
                                    "payoutMethod": matched_db_rec.get("payout_method", 2),
                                    "monthFrom": file_fm,
                                    "monthTo": expected_db_tm,
                                    "feeAmount": float(matched_db_rec.get("fee_amount") or 0.0),
                                    "premiumAmount": float(matched_db_rec.get("premium_amount") or 0.0),
                                    "commissionSplit": merged_splits,
                                    "_diff_text": diff_text
                                }
                                mismatch_policy_payloads.append(mismatch_payload)

        if group_idx % 50 == 0 or group_idx == total_groups:
            pct = (group_idx / total_groups) * 100.0
            sys.stdout.write(
                f"\r ⚡ [Stage 2 Chunked DB Verification] {group_idx:,}/{total_groups:,} policy groups verified ({pct:.1f}%) | "
                f"Matched Rows: {matched_count:,} | Unmatched Rows: {unmatched_count:,}"
            )
            sys.stdout.flush()

    print(f"\n   ✅ Stage 2 Database Verification Completed! {matched_count:,} matched rows, {unmatched_count:,} unmatched rows.")

    df_out["obtained_product_uuid"] = obtained_product_uuids
    df_out["obtained_policy_uuid"] = obtained_policy_uuids
    df_out["Stage2_Status"] = stage2_statuses
    df_out["Stage2_Unmatched_Reason"] = unmatched_reasons
    df_out["DB_Value"] = db_vals
    df_out["File_Value"] = file_vals

    matched_mask = df_out["Stage2_Status"] == "MATCHED"
    ambiguous_mask = df_out["Stage2_Status"] == "AMBIGUOUS_SIMILAR_MATCH"

    df_matched = df_out[matched_mask].copy()
    df_ambiguous = df_out[ambiguous_mask].copy()
    df_unmatched = df_out[~(matched_mask | ambiguous_mask)].copy()

    summary = {
        "stage2_processed": len(df_valid),
        "matched_records": len(df_matched),
        "ambiguous_records": len(df_ambiguous),
        "unmatched_records": len(df_unmatched),
    }

    # Dedup the missing payloads by policyUuid
    unique_missing_payloads = []
    seen_missing = set()
    for mp in missing_policy_payloads:
        if mp["policyUuid"] not in seen_missing:
            seen_missing.add(mp["policyUuid"])
            unique_missing_payloads.append(mp)

    # Dedup the mismatch payloads by rule_uuid
    unique_mismatch_payloads = []
    seen_mismatches = set()
    for mp in mismatch_policy_payloads:
        if mp["rule_uuid"] not in seen_mismatches:
            seen_mismatches.add(mp["rule_uuid"])
            unique_mismatch_payloads.append(mp)

    return df_matched, df_ambiguous, df_unmatched, unique_missing_payloads, unique_mismatch_payloads, invalid_user_uuid_payloads


# ---------------------------------------------------------------------------
# Output Export & Main CLI Pipeline
# ---------------------------------------------------------------------------
def export_phase1_reports(df_valid: pd.DataFrame, df_invalid: pd.DataFrame, versioned_dir: str) -> Tuple[str, str]:
    """Exports Phase 1 First Validation Reports."""
    os.makedirs(versioned_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    clean_path = os.path.join(versioned_dir, f"file_validated_clean_{timestamp}.xlsx")
    errors_path = os.path.join(versioned_dir, f"file_validated_errors_{timestamp}.xlsx")

    drop_cols = [c for c in df_valid.columns if str(c).startswith("_")]
    df_valid_clean = df_valid.drop(columns=drop_cols, errors="ignore")
    df_invalid_clean = df_invalid.drop(columns=drop_cols, errors="ignore")

    df_valid_clean.to_excel(clean_path, index=False)
    df_invalid_clean.to_excel(errors_path, index=False)

    return clean_path, errors_path


def export_phase2_reports(df_matched: pd.DataFrame, df_unmatched: pd.DataFrame, versioned_dir: str) -> Tuple[str, str]:
    """Exports Phase 2 DB Validated Reports."""
    os.makedirs(versioned_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    matched_path = os.path.join(versioned_dir, f"db_matched_commission_{timestamp}.xlsx")
    unmatched_path = os.path.join(versioned_dir, f"db_unmatched_commission_{timestamp}.xlsx")

    df_matched.to_excel(matched_path, index=False)
    df_unmatched.to_excel(unmatched_path, index=False)

    return matched_path, unmatched_path


def main():
    parser = argparse.ArgumentParser(description="2-Stage Policy Commission Import Validation & DB Verification Script")
    parser.add_argument("--file", type=str, help="Path to input APL policy commission upload file (CSV or Excel)")
    parser.add_argument("--rate-sheet", type=str, help="Path to master ProductWise Commission Rate.xlsx file")
    parser.add_argument("--norm-file", type=str, help="Path to product normalization table")
    parser.add_argument("--agent-roster", type=str, help="Path to Master Agent Roster file")
    parser.add_argument("--policy-mapping", type=str, help="Path to Recovered Policies Excel mapping file")
    parser.add_argument("--env-file", type=str, default=".env", help="Path to .env configuration file")
    parser.add_argument("--output-dir", type=str, default="pipeline_output/commission_validation_db", help="Directory to save output reports")
    parser.add_argument("--agency-id", type=int, default=5, help="Agency ID for database query filtering (default: 5)")
    parser.add_argument("--dry-run", action="store_true", help="Simulate DB queries for testing without active DB connection")
    parser.add_argument("--tolerance", type=float, default=0.02, help="Tolerance for commission value matching (default: 0.02 / +-0.02%%)")

    args = parser.parse_args()

    env_loaded = load_env_file(args.env_file)
    print("==========================================================")
    print("   PIEQ 2-STAGE POLICY COMMISSION VALIDATION & DB ENGINE  ")
    print("==========================================================")
    print(f"⚙️ Loaded {env_loaded} environment variable(s) from '{args.env_file}'.")

    agency_id = int(os.getenv("AGENCY_ID", args.agency_id))
    input_file = args.file or os.getenv("UPLOAD_FILE_PATH")

    if not input_file or not os.path.exists(input_file):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        possible = [f for f in os.listdir(script_dir) if f.endswith((".csv", ".xlsx")) and "commission" in f.lower()]
        if possible:
            input_file = os.path.join(script_dir, possible[0])

    if not input_file or not os.path.exists(input_file):
        print(f"❌ Error: Input file not specified or file not found. Provide via --file or UPLOAD_FILE_PATH in .env.")
        sys.exit(1)

    print(f"📖 Input Upload File: '{input_file}'")
    print(f"🏢 Target Agency ID: {agency_id}")

    print("\n----------------------------------------------------------")
    print("PRE-FLIGHT DATABASE CONNECTION CHECK")
    print("----------------------------------------------------------")
    db_mgr = DatabaseManager()
    db_connected = False
    if not args.dry_run:
        print("🔌 Testing Database Connection before processing...")
        db_connected = db_mgr.connect()
        if db_connected:
            print("   ✅ Pre-flight Database Connection: SUCCESSFUL!")
        else:
            print("   ⚠️ Pre-flight Database Connection: FAILED (Stage 2 will run in DRY-RUN mode).")
    else:
        print("   ℹ️ Running in DRY-RUN mode (Skipping Database connection test).")

    df_rules = None
    rate_sheet_path = args.rate_sheet or os.getenv("RATE_SHEET_PATH")
    if rate_sheet_path and os.path.exists(rate_sheet_path):
        print(f"\n📖 Loading Master Rate Sheet from '{rate_sheet_path}'...")
        df_rules = unpivot_product_commission_rates(rate_sheet_path)
        print(f"   ✅ Extracted {len(df_rules):,} active rate rules.")

    norm_dict = {}
    norm_file_path = args.norm_file or os.getenv("NORM_FILE_PATH")
    if norm_file_path and os.path.exists(norm_file_path):
        print(f"📖 Loading Product Normalization Map from '{norm_file_path}'...")
        norm_dict = load_product_normalization(norm_file_path)
        print(f"   ✅ Extracted {len(norm_dict):,} product mappings.")

    agent_roster_dict = {}
    agent_roster_path = args.agent_roster or os.getenv("AGENT_ROSTER_PATH")
    if agent_roster_path and os.path.exists(agent_roster_path):
        print(f"📖 Loading Agent Roster from '{agent_roster_path}'...")
        agent_roster_dict = load_agent_roster(agent_roster_path)
        print(f"   ✅ Extracted {len(agent_roster_dict):,} agent mappings.")

    print("\n----------------------------------------------------------")
    print("STAGE 1: DIFFERENTIAL ROLL-UP, SPLIT NORMALIZATION & RATE CHECK")
    print("----------------------------------------------------------")
    df_comm = load_file(input_file)
    print(f"📄 Loaded {len(df_comm):,} total input rows.")
    
    # Preprocess Policy Numbers to remove single quotes
    pol_col = next((c for c in df_comm.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    if pol_col:
        print(f"🧹 Preprocessing: Stripping single quotes from '{pol_col}'...")
        df_comm[pol_col] = df_comm[pol_col].astype(str).str.replace("'", "", regex=False)
        
    df_comm["_is_mapped"] = False
    
    # Optional Policy Mapping Replacement
    policy_mapping_path = args.policy_mapping or os.getenv("POLICY_MAPPING_PATH")
    if policy_mapping_path and os.path.exists(policy_mapping_path) and pol_col:
        try:
            print(f"\n📂 Found Policy Mapping file: '{policy_mapping_path}'")
            user_input = input("   Do you want to apply Processed Policy Numbers? (y/n): ").strip().lower()
            if user_input in ('y', 'yes'):
                print("   🔄 Loading mapping file and applying replacements...")
                df_map = pd.read_excel(policy_mapping_path) if policy_mapping_path.endswith((".xlsx", ".xls")) else pd.read_csv(policy_mapping_path)
                
                map_pol_col = next((c for c in df_map.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno")), None)
                map_proc_col = next((c for c in df_map.columns if str(c).strip().lower() in ("processed policy number", "processed_policy_number")), None)
                
                if map_pol_col and map_proc_col:
                    df_map[map_pol_col] = df_map[map_pol_col].astype(str).str.replace("'", "", regex=False).str.strip()
                    df_map[map_proc_col] = df_map[map_proc_col].astype(str).str.replace("'", "", regex=False).str.strip()
                    
                    df_map_valid = df_map[df_map[map_proc_col] != "nan"]
                    df_map_valid = df_map_valid[df_map_valid[map_proc_col] != ""]
                    
                    policy_dict = dict(zip(df_map_valid[map_pol_col], df_map_valid[map_proc_col]))
                    
                    df_comm["_mapped_pol"] = df_comm[pol_col].map(policy_dict)
                    mapped_mask = df_comm["_mapped_pol"].notna()
                    replaced_count = mapped_mask.sum()
                    
                    if replaced_count > 0:
                        df_comm.loc[mapped_mask, pol_col] = df_comm.loc[mapped_mask, "_mapped_pol"]
                        df_comm.loc[mapped_mask, "_is_mapped"] = True
                        print(f"   ✅ Successfully replaced {replaced_count:,} policy numbers using the mapping file.")
                    else:
                        print("   ⚠️ No matching policy numbers found in the mapping file to replace.")
                        
                    df_comm = df_comm.drop(columns=["_mapped_pol"], errors="ignore")
                else:
                    print(f"   ⚠️ Could not find 'Policy No' or 'Processed Policy Number' columns in mapping file. Skipping mapping.")
            else:
                print("   ⏩ Skipping Policy Mapping as per user input.")
        except Exception as e:
            print(f"   ❌ Error applying Policy Mapping: {e}")
            
    # Preprocessing: Upfront Deduplication
    ignore_cols = ["transaction_id", "source_batch_id", "uploaded_at", "id"]
    subset_cols = [c for c in df_comm.columns if str(c).strip().lower() not in ignore_cols]
    
    is_dup = df_comm.duplicated(subset=subset_cols, keep="first")
    df_duplicates = df_comm[is_dup].copy()
    df_comm_dedup = df_comm[~is_dup].copy().reset_index(drop=True)
    
    if len(df_duplicates) > 0:
        print(f"🧹 Preprocessing: Removed {len(df_duplicates):,} exact duplicate rows.")

    df_stage1_valid, df_stage1_invalid, df_stage1_zero, s1_summary = stage1_pre_validate_dataset(
        df_comm=df_comm_dedup,
        df_rules=df_rules,
        norm_dict=norm_dict,
        agent_roster_dict=agent_roster_dict,
        tolerance=args.tolerance,
    )

    print(f"✅ Stage 1 Pre-Validation Passed : {s1_summary['stage1_valid_records']:,} rows")
    print(f"❌ Stage 1 Pre-Validation Failed : {s1_summary['stage1_invalid_records']:,} rows (Bypassing DB queries)")
    if s1_summary['stage1_zero_ignored'] > 0:
        print(f"🟡 Stage 1 Zero Comm Ignored   : {s1_summary['stage1_zero_ignored']:,} rows")

    # Create a single dynamic version folder for this run
    run_version = f"v{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    versioned_dir = os.path.join(args.output_dir, run_version)
    input_dir = os.path.join(versioned_dir, "input")
    output_dir = os.path.join(versioned_dir, "output")
    
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    import shutil
    shutil.copy2(input_file, os.path.join(input_dir, os.path.basename(input_file)))

    # Phase 1 Exports
    phase1_clean_path, phase1_error_path = export_phase1_reports(df_stage1_valid, df_stage1_invalid, output_dir)
    print(f" 📁 Phase 1 Clean File  : {phase1_clean_path}")
    print(f" 📁 Phase 1 Error File  : {phase1_error_path}")
    
    dup_path = ""
    if len(df_duplicates) > 0:
        dup_path = os.path.join(output_dir, f"file_validated_duplicates_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        df_duplicates.to_excel(dup_path, index=False)
        print(f" 📁 Phase 1 Duplicates  : {dup_path}")
        
    zero_path = ""
    if len(df_stage1_zero) > 0:
        zero_path = os.path.join(output_dir, f"file_validated_zero_comm_ignored_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        df_stage1_zero.to_excel(zero_path, index=False)
        print(f" 📁 Phase 1 Zero Comm   : {zero_path}")

    print("\n----------------------------------------------------------")
    print(f"STAGE 2: READ-ONLY DATABASE VERIFICATION (AGENCY_ID = {agency_id})")
    print("----------------------------------------------------------")

    dry_run_mode = args.dry_run or (not db_connected)
    if dry_run_mode and not args.dry_run:
        print("⚠️ Notice: Database unestablished. Executing Stage 2 in DRY-RUN simulation mode.")

    df_stage2_matched, df_stage2_ambiguous, df_stage2_unmatched, unique_missing_payloads, unique_mismatch_payloads, invalid_user_uuid_payloads = stage2_db_verify_dataset(
        df_valid=df_stage1_valid,
        db_mgr=db_mgr,
        agency_id=agency_id,
        dry_run=dry_run_mode,
        tolerance=args.tolerance,
    )
    db_mgr.close()

    if invalid_user_uuid_payloads:
        missing_uuid_file = os.path.join(output_dir, f"rejected_missing_user_uuid_payloads_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        with open(missing_uuid_file, "w") as f:
            json.dump(invalid_user_uuid_payloads, f, indent=2)
        print(f" ⚠️ Excluded {len(invalid_user_uuid_payloads):,} payload(s) with missing user_uuid. Logged to: {missing_uuid_file}")

    df_all_unmatched = pd.concat([df_stage1_invalid, df_stage2_unmatched], ignore_index=True)

    # Phase 2 Exports
    phase2_matched_path, phase2_unmatched_path = export_phase2_reports(df_stage2_matched, df_stage2_unmatched, output_dir)
    
    ambiguous_path = ""
    if len(df_stage2_ambiguous) > 0:
        ambiguous_path = os.path.join(output_dir, f"file_validated_ambiguous_matches_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        df_stage2_ambiguous.to_excel(ambiguous_path, index=False)
        print(f" 📁 Phase 2 Ambiguous  : {ambiguous_path}")

    tot_records = len(df_comm)
    s1_dup = len(df_duplicates)
    tot_unique = len(df_comm_dedup)
    
    s1_zero = len(df_stage1_zero)
    s1_clean = len(df_stage1_valid)
    s1_err = len(df_stage1_invalid)
    
    s2_match = len(df_stage2_matched)
    s2_ambig = len(df_stage2_ambiguous)
    s2_unmatch = len(df_stage2_unmatched)
    tot_unmatch = len(df_all_unmatched)
    
    pol_col = next((c for c in df_comm.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    
    def count_unique_pols(df):
        if len(df) == 0 or not pol_col or pol_col not in df.columns: return 0
        return df[pol_col].nunique()
        
    s1_err_pols = count_unique_pols(df_stage1_invalid)
    s2_unmatch_pols = count_unique_pols(df_stage2_unmatched)
    s2_ambig_pols = count_unique_pols(df_stage2_ambiguous)
    tot_unmatch_pols = count_unique_pols(df_all_unmatched)
    
    s1_err_str = f"{s1_err:,} ({s1_err_pols:,} pols)" if s1_err > 0 else "0"
    s2_unmatch_str = f"{s2_unmatch:,} ({s2_unmatch_pols:,} pols)" if s2_unmatch > 0 else "0"
    s2_ambig_str = f"{s2_ambig:,} ({s2_ambig_pols:,} pols)" if s2_ambig > 0 else "0"
    tot_unmatch_str = f"{tot_unmatch:,} ({tot_unmatch_pols:,} pols)" if tot_unmatch > 0 else "0"

    pct_s1_dup = round((s1_dup / tot_records) * 100, 2) if tot_records > 0 else 0
    pct_tot_unique = round((tot_unique / tot_records) * 100, 2) if tot_records > 0 else 0
    
    pct_s1_zero = round((s1_zero / tot_unique) * 100, 2) if tot_unique > 0 else 0
    pct_s1_clean = round((s1_clean / tot_unique) * 100, 2) if tot_unique > 0 else 0
    pct_s1_err = round((s1_err / tot_unique) * 100, 2) if tot_unique > 0 else 0
    
    pct_s2_match = round((s2_match / s1_clean) * 100, 2) if s1_clean > 0 else 0
    pct_s2_ambig = round((s2_ambig / s1_clean) * 100, 2) if s1_clean > 0 else 0
    pct_s2_unmatch = round((s2_unmatch / s1_clean) * 100, 2) if s1_clean > 0 else 0
    
    pct_tot_unmatch = round((tot_unmatch / tot_unique) * 100, 2) if tot_unique > 0 else 0

    print("\n=========================================================================================")
    print("                              FINAL EXECUTION SUMMARY TABLE                              ")
    print("=========================================================================================")
    print("+------------------------------------+-----------------------+------------+------------+")
    print("| Metric / Pipeline Stage            | Record Count          | Out of     | Percentage |")
    print("+------------------------------------+-----------------------+------------+------------+")
    print(f"| Total Input Records Uploaded       | {tot_records:<21,} | {tot_records:<10,} | 100.00%    |")
    if s1_dup > 0:
        print(f"| 🟡 Exact Duplicates (Removed)      | {s1_dup:<21,} | {tot_records:<10,} | {pct_s1_dup:>6.2f}%   |")
    print(f"| 🔵 Unique Records Processed        | {tot_unique:<21,} | {tot_records:<10,} | {pct_tot_unique:>6.2f}%   |")
    print("+------------------------------------+-----------------------+------------+------------+")
    if s1_zero > 0:
        print(f"| 🟡 Phase 1 Zero Comm (Ignored)     | {s1_zero:<21,} | {tot_unique:<10,} | {pct_s1_zero:>6.2f}%   |")
    print(f"| 🟢 Phase 1 Pre-Validated Clean     | {s1_clean:<21,} | {tot_unique:<10,} | {pct_s1_clean:>6.2f}%   |")
    print(f"| 🔴 Phase 1 Pre-Validation Errors   | {s1_err_str:<21} | {tot_unique:<10,} | {pct_s1_err:>6.2f}%   |")
    print("+------------------------------------+-----------------------+------------+------------+")
    print(f"| 🟢 Phase 2 DB Matched Records      | {s2_match:<21,} | {s1_clean:<10,} | {pct_s2_match:>6.2f}%   |")
    if s2_ambig > 0:
        print(f"| 🟡 Phase 2 Ambiguous (Skipped)     | {s2_ambig_str:<21} | {s1_clean:<10,} | {pct_s2_ambig:>6.2f}%   |")
    print(f"| 🔴 Phase 2 DB Unmatched Records    | {s2_unmatch_str:<21} | {s1_clean:<10,} | {pct_s2_unmatch:>6.2f}%   |")
    print("+------------------------------------+-----------------------+------------+------------+")
    print(f"| ❌ TOTAL FAILED / UNMATCHED (P1+P2)| {tot_unmatch_str:<21} | {tot_unique:<10,} | {pct_tot_unmatch:>6.2f}%   |")
    print("+------------------------------------+-----------------------+------------+------------+")
    
    if s1_err > 0 or s2_unmatch > 0 or s2_ambig > 0:
        print("\n-----------------------------------------------------------------------------------------")
        print("                              ERROR & MISMATCH BREAKDOWN                                 ")
        print("-----------------------------------------------------------------------------------------")
        
        def group_error_reason(reason: str) -> str:
            r = str(reason).lower()
            if "not found in policies table" in r:
                return "Policy number not found in policies table"
            if "policy group not found in db at all" in r or "no commission rule record found in policy_commission table" in r or "not found in db for" in r or "not found in db policy commission splits" in r or "not found in db commiss" in r:
                return "No commission rule record found in policy_commission table"
            if "found in db, but db commission_value is" in r:
                return "Commission mismatch with DB"
            if "similar match for" in r and "returned multiple different policies" in r:
                return "Ambiguous match (multiple policies)"
            if "split rows instead of expected" in r:
                return "Invalid split row count"
            if "not found in master rate sheet" in r:
                return "Product missing in rate sheet"
            if "out of rate tier ranges" in r:
                return "Month out of rate tier range"
            if "group split share sum is" in r:
                return "Split share sum is not 100%"
            if "carrier commission mismatch" in r:
                return "Carrier commission mismatch"
            return str(reason)

        if s1_err > 0 and "Stage1_Error_Reason" in df_stage1_invalid.columns:
            print("🔴 Phase 1 Pre-Validation Errors:")
            s1_counts = df_stage1_invalid["Stage1_Error_Reason"].apply(group_error_reason).value_counts()
            for reason, count in s1_counts.items():
                print(f"   - {str(reason)[:70]:<70} : {count} records")
                
        if s2_unmatch > 0 and "Stage2_Unmatched_Reason" in df_stage2_unmatched.columns:
            print("\n🔴 Phase 2 DB Unmatched:")
            s2_counts = df_stage2_unmatched["Stage2_Unmatched_Reason"].apply(group_error_reason).value_counts()
            for reason, count in s2_counts.items():
                print(f"   - {str(reason)[:70]:<70} : {count} records")
                
        if s2_ambig > 0 and "Stage2_Unmatched_Reason" in df_stage2_ambiguous.columns:
            print("\n🟡 Phase 2 Ambiguous:")
            s2a_counts = df_stage2_ambiguous["Stage2_Unmatched_Reason"].apply(group_error_reason).value_counts()
            for reason, count in s2a_counts.items():
                print(f"   - {str(reason)[:70]:<70} : {count} records")

    print("\n-----------------------------------------------------------------------------------------")
    print("OUTPUT FILE REPORTS GENERATED:")
    print(f" 📁 Phase 1 Clean File  : {phase1_clean_path}")
    print(f" 📁 Phase 1 Error File  : {phase1_error_path}")
    if s1_dup > 0:
        print(f" 📁 Phase 1 Duplicates  : {dup_path}")
    if s1_zero > 0:
        print(f" 📁 Phase 1 Zero Comm   : {zero_path}")
    print(f" 📁 Phase 2 Matched File: {phase2_matched_path}")
    if s2_ambig > 0:
        print(f" 📁 Phase 2 Ambiguous   : {ambiguous_path}")
    print(f" 📁 Phase 2 Unmatched   : {phase2_unmatched_path}")
    print("=========================================================================================")
    # Write comprehensive markdown summary report
    # Prepare detailed mismatches (if any)
    detail_rows = None
    if not df_stage2_unmatched.empty:
        # Detect actual column names dynamically
        pol_col = next((c for c in df_stage2_unmatched.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
        prod_col = next((c for c in df_stage2_unmatched.columns if str(c).strip().lower() in ("normalized_product_name", "product name", "product")), None)
        fm_col = next((c for c in df_stage2_unmatched.columns if str(c).strip().lower() in ("from month", "from_month", "month_from")), None)
        tm_col = next((c for c in df_stage2_unmatched.columns if str(c).strip().lower() in ("to month", "to_month", "month_to")), None)
        detail_cols = [c for c in [pol_col, prod_col, fm_col, tm_col, "DB_Value", "File_Value", "Stage2_Unmatched_Reason"] if c and c in df_stage2_unmatched.columns]
        detail_rows = df_stage2_unmatched[detail_cols].fillna("") if detail_cols else None
    # Create markdown summary file
    md_summary_path = os.path.join(output_dir, f"validation_summary_{run_version}.md")
    try:
        with open(md_summary_path, "w") as md_file:
            md_file.write("# Validation Summary Report\n\n")
            md_file.write(f"**Total Input Records Uploaded:** {tot_records}\n")
            if s1_dup > 0:
                md_file.write(f"**Exact Duplicates (Removed):** {s1_dup} ({pct_s1_dup:.2f}%)\n")
            md_file.write(f"**Unique Records Processed:** {tot_unique} ({pct_tot_unique:.2f}%)\n\n")
            if s1_zero > 0:
                md_file.write(f"**Phase 1 Zero Comm (Ignored):** {s1_zero} ({pct_s1_zero:.2f}%)\n")
            md_file.write(f"**Phase 1 Clean:** {s1_clean} ({pct_s1_clean:.2f}%)\n")
            md_file.write(f"**Phase 1 Errors:** {s1_err} ({pct_s1_err:.2f}%)\n\n")
            md_file.write(f"**Phase 2 Matched:** {s2_match} ({pct_s2_match:.2f}%)\n")
            if s2_ambig > 0:
                md_file.write(f"**Phase 2 Ambiguous (Skipped):** {s2_ambig} ({pct_s2_ambig:.2f}%)\n")
            md_file.write(f"**Phase 2 Unmatched:** {s2_unmatch} ({pct_s2_unmatch:.2f}%)\n\n")
            md_file.write(f"**TOTAL FAILED / UNMATCHED (P1+P2):** {tot_unmatch} ({pct_tot_unmatch:.2f}%)\n\n")
            if detail_rows is not None and not detail_rows.empty:
                md_file.write("## Detailed Stage 2 Mismatches\n\n")
                for _, row in detail_rows.iterrows():
                    policy_val = str(row.get(pol_col, "")) if pol_col and pol_col in detail_rows.columns else ""
                    product_val = str(row.get(prod_col, "")) if prod_col and prod_col in detail_rows.columns else ""
                    db_val = str(row.get("DB_Value", "")) if "DB_Value" in detail_rows.columns else ""
                    file_val = str(row.get("File_Value", "")) if "File_Value" in detail_rows.columns else ""
                    reason = str(row.get("Stage2_Unmatched_Reason", "")) if "Stage2_Unmatched_Reason" in detail_rows.columns else ""
                    from_month = str(row.get(fm_col, "")).replace(".0", "") if fm_col and fm_col in detail_rows.columns else ""
                    to_month = str(row.get(tm_col, "")).replace(".0", "") if tm_col and tm_col in detail_rows.columns else ""
                    month_range = f" | **Month Range:** {from_month} → {to_month}" if from_month else ""

                    md_file.write("---\n\n")
                    md_file.write(f"**Policy No:** {policy_val} | **Product:** {product_val}{month_range}\n\n")

                    # Try to parse both as JSON and render comparison table
                    db_parsed = None
                    file_parsed = None
                    try:
                        if db_val and db_val not in ("NOT_FOUND", "COMMISSION_NOT_FOUND", ""):
                            db_parsed = json.loads(db_val)
                    except (json.JSONDecodeError, TypeError):
                        pass
                    try:
                        if file_val and file_val not in ("", "NOT_FOUND"):
                            file_parsed = json.loads(file_val)
                    except (json.JSONDecodeError, TypeError):
                        pass

                    if isinstance(db_parsed, list) and isinstance(file_parsed, list):
                        # Build lookup maps by order
                        db_by_order = {item.get("order"): item for item in db_parsed if isinstance(item, dict)}
                        file_by_order = {item.get("order"): item for item in file_parsed if isinstance(item, dict)}
                        all_orders = sorted(set(list(db_by_order.keys()) + list(file_by_order.keys())))

                        # Collect all rows and highlight mismatches
                        table_rows = []
                        has_any_mismatch = False
                        for o in all_orders:
                            db_entry = db_by_order.get(o, {})
                            file_entry = file_by_order.get(o, {})
                            
                            db_cv = db_entry.get("commission_value")
                            file_cv = file_entry.get("commission_value")
                            
                            db_agent = db_entry.get("user_name", "") or ""
                            file_agent = file_entry.get("user_name", "") or ""
                            db_code = db_entry.get("user_code", "") or ""
                            file_code = file_entry.get("user_code", "") or ""

                            # Check value match
                            val_match = db_cv is not None and file_cv is not None and abs(float(db_cv) - float(file_cv)) <= 0.05
                            
                            # Check agent match
                            db_norm = " ".join(str(db_agent).lower().split())
                            file_norm = " ".join(str(file_agent).lower().split())
                            agent_match = False
                            if file_code and db_code and str(file_code).strip() == str(db_code).strip():
                                agent_match = True
                            elif file_norm and db_norm and file_norm == db_norm:
                                agent_match = True
                            elif not file_norm and not db_norm:
                                agent_match = True  # Both missing agent (e.g. fixed fee)
                            elif str(o) == "1" and (
                                ("main line benefits" in file_norm and "mlb parallel run" in db_norm) or
                                ("mlb parallel run" in file_norm and "main line benefits" in db_norm)
                            ):
                                agent_match = True

                            role = file_entry.get("role") or db_entry.get("role") or ""

                            # Ignore Sub Agent N/A vs 0% mismatch
                            if "sub agent" in role.lower() and file_cv is None and db_cv is not None and float(db_cv) == 0.0:
                                val_match = True
                                agent_match = True

                            is_mismatch = not (val_match and agent_match)
                            if is_mismatch:
                                has_any_mismatch = True

                            file_agent_str = f"{file_agent} ({file_code})" if file_code else str(file_agent)
                            db_agent_str = f"{db_agent} ({db_code})" if db_code else str(db_agent)
                            
                            file_cv_str = f"{file_cv}%" if file_cv is not None else "_N/A_"
                            db_cv_str = f"{db_cv}%" if db_cv is not None else "_N/A_"
                            
                            # Highlight mismatching rows
                            if is_mismatch:
                                o_str = f"**{o} ❌**"
                                role = f"**{role}**"
                                file_agent_str = f"**{file_agent_str}**"
                                db_agent_str = f"**{db_agent_str}**"
                                file_cv_str = f"**{file_cv_str}**"
                                db_cv_str = f"**{db_cv_str}**"
                            else:
                                o_str = str(o)
                            
                            table_rows.append((o_str, role, file_agent_str, db_agent_str, file_cv_str, db_cv_str))

                        if has_any_mismatch and table_rows:
                            md_file.write(f"**Reason:** {reason}\n\n")
                            md_file.write("| Order | Role | File Agent (Code) | DB Agent (Code) | File Value | DB Value |\n")
                            md_file.write("|-------|------|-------------------|-----------------|------------|----------|\n")
                            for o_str, role, fa, da, fv, dv in table_rows:
                                md_file.write(f"| {o_str} | {role} | {fa} | {da} | {fv} | {dv} |\n")
                            md_file.write("\n")
                        else:
                            md_file.write(f"**Reason:** {reason}\n\n")
                    else:
                        # NOT_FOUND or non-JSON — just show the reason
                        md_file.write(f"**Reason:** {reason}\n\n")

                md_file.write("\n")
    except Exception as e:
        print(f"⚠️ Failed to write markdown summary: {e}")
    print(f"\n📝 Validation summary markdown saved to: {md_summary_path}")

    if unique_missing_payloads:
        print("\n" + "="*80)
        print(f"🤖 API AUTO-IMPORT: Found {len(unique_missing_payloads)} policies with completely missing commission rules in the DB.")
        print("="*80)
        choice = input("Would you like to auto-create them via the POST API? (y/n): ").strip().lower()
        if choice in ('y', 'yes'):
            import sys
            import logging
            import urllib.request
            import urllib.error
            from dotenv import load_dotenv

            other_script_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "policy_commission_import_validation_script"))
            other_env_path = os.path.join(other_script_dir, ".env")
            
            if os.path.exists(other_env_path):
                # Use override=False so we don't overwrite the corrected local .env credentials!
                load_dotenv(other_env_path, override=False)
                
            if other_script_dir not in sys.path:
                sys.path.append(other_script_dir)
            try:
                from import_policy_commission import fetch_sso_token, confirm_payload
                
                logger = logging.getLogger("AutoImport")
                logger.setLevel(logging.INFO)
                if not logger.handlers:
                    ch = logging.StreamHandler()
                    logger.addHandler(ch)

                api_base = os.environ.get("API_BASE_URL", "https://prod.api.pieq.ai").rstrip("/")
                sso_url_raw = os.environ.get("AUTH_URL", "https://prod.auth.pieq.ai")
                sso_url = f"{sso_url_raw.rstrip('/')}/realms/pieq-sso/protocol/openid-connect/token"
                client_id = os.environ.get("KEYCLOAK_CLIENT_ID", "pieq-app")
                username = os.environ.get("USERNAME")
                password = os.environ.get("PASSWORD")
                
                token = fetch_sso_token(logger, sso_url, client_id, username, password)
                
                if token:
                    auto_confirm = False
                    for idx, payload in enumerate(unique_missing_payloads):
                        action = confirm_payload(logger, idx+1, len(unique_missing_payloads), payload, auto_confirm)
                        if action == "quit":
                            break
                        elif action == "all":
                            auto_confirm = True
                            action = "yes"
                        elif action == "skip":
                            continue

                        if action == "yes":
                            payload_to_send = payload.copy()
                            post_url = f"{api_base}/api/v1/policy/commission"
                            headers = {
                                "Content-Type": "application/json",
                                "Authorization": f"Bearer {token}"
                            }
                            if "X_CLIENT_ID" in os.environ:
                                headers["x-client-id"] = os.environ["X_CLIENT_ID"]
                            if "X_SERVICE_NAME" in os.environ:
                                headers["x-service-name"] = os.environ["X_SERVICE_NAME"]
                                
                            # pop policyNumber because it's not part of PolicyCommissionRequest
                            pol_num = payload_to_send.pop("policyNumber", "Unknown")
                            
                            # The POST endpoint accepts a List<PolicyCommissionRequest>
                            req = urllib.request.Request(
                                post_url, 
                                data=json.dumps([payload_to_send]).encode("utf-8"), 
                                headers=headers, 
                                method="POST"
                            )
                            try:
                                with urllib.request.urlopen(req, timeout=30.0) as resp:
                                    res_bytes = resp.read()
                                    res_json = json.loads(res_bytes.decode("utf-8"))
                                    logger.info(f"✅ Success POSTing {pol_num} to {post_url}: {res_json}")
                            except urllib.error.HTTPError as e:
                                if e.code == 401:
                                    logger.warning(f"⚠️ Token expired (HTTP 401). Re-fetching token and retrying POST for {pol_num}...")
                                    token = fetch_sso_token(logger, sso_url, client_id, username, password)
                                    if token:
                                        headers["Authorization"] = f"Bearer {token}"
                                        req = urllib.request.Request(
                                            post_url, 
                                            data=json.dumps([payload_to_send]).encode("utf-8"), 
                                            headers=headers, 
                                            method="POST"
                                        )
                                        try:
                                            with urllib.request.urlopen(req, timeout=30.0) as resp:
                                                res_bytes = resp.read()
                                                res_json = json.loads(res_bytes.decode("utf-8"))
                                                logger.info(f"✅ Success POSTing {pol_num} to {post_url} after token refresh: {res_json}")
                                                continue
                                        except urllib.error.HTTPError as retry_e:
                                            e = retry_e
                                        except Exception as retry_err:
                                            logger.error(f"❌ Failed to POST {pol_num} on retry: {retry_err}")
                                            continue
                                err_body = e.read().decode("utf-8", errors="ignore") if hasattr(e, 'read') else str(e)
                                logger.error(f"❌ Failed to POST {pol_num} to {post_url} (HTTP {getattr(e, 'code', 'Error')}): {err_body}")
                            except Exception as err:
                                logger.error(f"❌ Failed to POST {pol_num}: {err}")
                else:
                    print("⚠️ Could not generate API token. Make sure .env credentials are set.")
            except ImportError as ie:
                print(f"⚠️ Could not import API logic from {other_script_dir}. ({ie})")
            except Exception as e:
                print(f"⚠️ API Auto-Import failed: {e}")

    if unique_mismatch_payloads:
        print("\n" + "="*80)
        print(f"🤖 API AUTO-UPDATE: Found {len(unique_mismatch_payloads)} mismatched commission rules.")
        print("="*80)
        choice = input("Would you like to auto-update them via the PUT API? (y/n): ").strip().lower()
        if choice in ('y', 'yes'):
            import sys
            import logging
            import urllib.request
            import urllib.error
            from dotenv import load_dotenv

            other_script_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "policy_commission_import_validation_script"))
            other_env_path = os.path.join(other_script_dir, ".env")
            if os.path.exists(other_env_path):
                # Use override=False so we don't overwrite the corrected local .env credentials!
                load_dotenv(other_env_path, override=False)
                
            if other_script_dir not in sys.path:
                sys.path.append(other_script_dir)
            try:
                from import_policy_commission import fetch_sso_token, confirm_payload
                
                logger = logging.getLogger("AutoUpdate")
                logger.setLevel(logging.INFO)
                if not logger.handlers:
                    ch = logging.StreamHandler()
                    logger.addHandler(ch)

                api_base = os.environ.get("API_BASE_URL", "https://prod.api.pieq.ai").rstrip("/")
                sso_url_raw = os.environ.get("AUTH_URL", "https://prod.auth.pieq.ai")
                sso_url = f"{sso_url_raw.rstrip('/')}/realms/pieq-sso/protocol/openid-connect/token"
                client_id = os.environ.get("KEYCLOAK_CLIENT_ID", "pieq-app")
                username = os.environ.get("USERNAME")
                password = os.environ.get("PASSWORD")
                
                token = fetch_sso_token(logger, sso_url, client_id, username, password)
                
                if token:
                    auto_confirm = False
                    put_failures_path = os.path.join(output_dir, f"put_api_failures_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
                    
                    for idx, payload in enumerate(unique_mismatch_payloads):
                        payload_to_send = payload.copy()
                        diff_text = payload_to_send.pop("_diff_text", "")
                        action = confirm_payload(logger, idx+1, len(unique_mismatch_payloads), payload_to_send, auto_confirm, action_name="PUT", diff_text=diff_text)
                        if action == "quit":
                            break
                        elif action == "all":
                            auto_confirm = True
                            action = "yes"
                        elif action == "skip":
                            continue

                        if action == "yes":
                            rule_uuid = payload_to_send.pop("rule_uuid") # Extract and remove from body
                            pol_num = payload_to_send.pop("policyNumber", "Unknown") # Remove helper field before sending
                            put_url = f"{api_base}/api/v1/policy/commission/{rule_uuid}"
                            headers = {
                                "Content-Type": "application/json",
                                "Authorization": f"Bearer {token}"
                            }
                            if "X_CLIENT_ID" in os.environ:
                                headers["x-client-id"] = os.environ["X_CLIENT_ID"]
                            if "X_SERVICE_NAME" in os.environ:
                                headers["x-service-name"] = os.environ["X_SERVICE_NAME"]
                                
                            req = urllib.request.Request(
                                put_url, 
                                data=json.dumps(payload_to_send).encode("utf-8"), 
                                headers=headers, 
                                method="PUT"
                            )
                            try:
                                with urllib.request.urlopen(req, timeout=30.0) as resp:
                                    res_bytes = resp.read()
                                    res_json = json.loads(res_bytes.decode("utf-8"))
                                    logger.info(f"✅ Success PUT updating {pol_num} at {put_url}")
                            except urllib.error.HTTPError as e:
                                if e.code == 401:
                                    logger.warning(f"⚠️ Token expired (HTTP 401). Re-fetching token and retrying PUT for {pol_num}...")
                                    token = fetch_sso_token(logger, sso_url, client_id, username, password)
                                    if token:
                                        headers["Authorization"] = f"Bearer {token}"
                                        req = urllib.request.Request(
                                            put_url, 
                                            data=json.dumps(payload_to_send).encode("utf-8"), 
                                            headers=headers, 
                                            method="PUT"
                                        )
                                        try:
                                            with urllib.request.urlopen(req, timeout=30.0) as resp:
                                                res_bytes = resp.read()
                                                res_json = json.loads(res_bytes.decode("utf-8"))
                                                logger.info(f"✅ Success PUT updating {pol_num} at {put_url} after token refresh")
                                                continue
                                        except urllib.error.HTTPError as retry_e:
                                            e = retry_e
                                        except Exception as retry_err:
                                            logger.error(f"❌ Failed to PUT {pol_num} on retry: {retry_err}")
                                            continue
                                err_body = e.read().decode("utf-8", errors="ignore") if hasattr(e, 'read') else str(e)
                                logger.error(f"❌ Failed to PUT {pol_num} to {put_url} (HTTP {getattr(e, 'code', 'Error')}): {err_body}")
                                with open(put_failures_path, "a") as f:
                                    f.write(f"[{datetime.datetime.now().isoformat()}] HTTP {getattr(e, 'code', 'Error')} for {pol_num} ({put_url}): {err_body}\n")
                            except Exception as err:
                                logger.error(f"❌ Failed to PUT {pol_num}: {err}")
                                with open(put_failures_path, "a") as f:
                                    f.write(f"[{datetime.datetime.now().isoformat()}] Exception for {pol_num} ({put_url}): {err}\n")
                                    
                    if os.path.exists(put_failures_path):
                        print(f"\n⚠️ Some PUT updates failed. Check logs at: {put_failures_path}")
                else:
                    print("⚠️ Could not generate API token. Make sure .env credentials are set.")
            except ImportError as ie:
                print(f"⚠️ Could not import API logic from {other_script_dir}. ({ie})")
            except Exception as e:
                print(f"⚠️ API Auto-Update failed: {e}")

if __name__ == "__main__":
    main()
