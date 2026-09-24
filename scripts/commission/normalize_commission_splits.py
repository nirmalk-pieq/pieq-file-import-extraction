#!/usr/bin/env python3
"""
Standalone Commission Roll-up & Split Normalization Utility

Reads policy commission datasets, groups rows by (Policy No, Product, From Month),
and performs sequence rate differential roll-ups and percentage split normalizations so that:
1. Sequence rates sum to exactly 100.00% split shares for each policy group.
2. Fixed Fee amounts are converted to incremental differential amounts.
3. Carrier Commission total pool values and Payout Methods are populated.
"""

import argparse
import math
import os
import sys
import time
import pandas as pd


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
    """Parses percentage strings or numeric values into float."""
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace("%", "").strip()
    if not val_str:
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0


def format_amount(val: float) -> str:
    """Formats float values to standard currency strings."""
    if val < 0:
        return f"-${abs(val):,.2f}"
    return f"${val:,.2f}"


def format_rate(val: float) -> str:
    """Formats float values into percentage strings (e.g. '40.00 %')."""
    return f"{val:.2f} %"


def clean_key(series: pd.Series) -> pd.Series:
    """Normalizes string keys for grouping."""
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.strip("'")
        .str.strip('"')
        .str.upper()
        .apply(lambda s: " ".join(s.split()))
    )


def rollup_commission_splits(df: pd.DataFrame) -> pd.DataFrame:
    """
    Performs sequence differential roll-up and percentage split normalization on DataFrame.
    Returns transformed DataFrame with normalized Rate (summing to 100%), Amount, Carrier Commission, and Payout Method.
    """
    df_out = df.copy()

    # Detect key columns
    pol_col = next((c for c in df_out.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno")), None)
    prod_col = next((c for c in df_out.columns if str(c).strip().lower() in ("product", "product name", "product_description")), None)
    fm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("from month", "from_month", "month")), None)
    seq_col = next((c for c in df_out.columns if str(c).strip().lower() in ("sequence", "seq")), None)
    rate_col = next((c for c in df_out.columns if str(c).strip().lower() == "rate"), None)
    amt_col = next((c for c in df_out.columns if str(c).strip().lower() in ("amount", "amt")), None)

    if not pol_col:
        raise KeyError(f"Required column 'Policy No' missing in dataset. Available: {list(df_out.columns)}")

    # Ensure output columns exist
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
    total_groups = len(grouped)
    print(f"🔄 Normalizing commission splits across {total_groups:,} policy group(s)...")

    res_amt = list(df_out[amt_col]) if amt_col else []
    res_rate = list(df_out[rate_col]) if rate_col else []
    res_carrier_comm = ["$0.00"] * len(df_out)
    res_payout_method = [""] * len(df_out)

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
            # CASE 1: FIXED FEE ($) — Roll up amounts & calculate 100% split percentages
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
                highest_seq_indices_in_group = [i for i, s in enumerate(seqs) if s == max_seq]
                primary_highest_idx = highest_seq_indices_in_group[-1]

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
                res_carrier_comm[global_idx] = carrier_comm_str
                res_payout_method[global_idx] = "FIXED FEE"

        else:
            # CASE 2: PERCENTAGE (%) — Compute rate differentials & normalize split shares to 100%
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
                highest_seq_indices_in_group = [i for i, s in enumerate(seqs) if s == max_seq]
                primary_highest_idx = highest_seq_indices_in_group[-1]

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
            for global_idx, r_val in zip(indices, rates_calc):
                if rate_col:
                    res_rate[global_idx] = format_rate(r_val)
                res_carrier_comm[global_idx] = carrier_comm_str
                res_payout_method[global_idx] = "PERCENTAGE"

    if amt_col:
        df_out[amt_col] = res_amt
    if rate_col:
        df_out[rate_col] = res_rate
    df_out["Carrier Commission"] = res_carrier_comm
    df_out["Payout Method"] = res_payout_method

    # Cleanup temporary helper columns
    df_out = df_out.drop(columns=["_amt_num", "_rate_num", "_seq_num", "_pol_clean", "_prod_clean", "_fm_clean"])
    return df_out


def main():
    parser = argparse.ArgumentParser(
        description="Standalone Commission Roll-up & Split Normalization Utility."
    )
    parser.add_argument(
        "-i", "--input",
        required=True,
        help="Path to input policy commission Excel or CSV file."
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Path to output Excel or CSV file (default: adds '_normalized' suffix to input path)."
    )

    args = parser.parse_args()

    input_path = args.input
    if not os.path.exists(input_path):
        print(f"❌ Error: Input file not found at '{input_path}'", file=sys.stderr)
        sys.exit(1)

    output_path = args.output
    if not output_path:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_normalized{ext}"

    start_time = time.time()
    print("==================================================")
    print("   COMMISSION ROLL-UP & SPLIT NORMALIZATION TOOL  ")
    print("==================================================")
    print(f"📖 Reading input file: '{input_path}'...")

    _, ext = os.path.splitext(input_path.lower())
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(input_path)
    else:
        df = pd.read_csv(input_path, low_memory=False)

    print(f"   ✅ Loaded {len(df):,} record(s).")

    df_normalized = rollup_commission_splits(df)

    print(f"\n💾 Saving normalized file to: '{output_path}'...")
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    _, out_ext = os.path.splitext(output_path.lower())
    if out_ext in (".xlsx", ".xls"):
        try:
            df_normalized.to_excel(output_path, index=False, engine="openpyxl")
        except Exception as ex:
            csv_path = output_path.rsplit(".", 1)[0] + ".csv"
            print(f"⚠️ Excel export failed ({ex}), saving as CSV instead to '{csv_path}'...")
            df_normalized.to_csv(csv_path, index=False)
            output_path = csv_path
    else:
        df_normalized.to_csv(output_path, index=False)

    elapsed = time.time() - start_time
    print(f"   ✅ Processed and saved {len(df_normalized):,} record(s) in {elapsed:.2f} seconds.")
    print("==================================================")


if __name__ == "__main__":
    main()
