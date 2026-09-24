#!/usr/bin/env python3
"""
Policy Product & Carrier Normalization Utility

Matches 'Product' in policy file(s) or folder(s) against a Product Normalization reference file.
Updates the 'Product' column with 'New Product Name' and updates/adds the 'Carrier' column.
Supports single or multiple target files/folders, fuzzy matching, and flexible column mappings.
"""

import os
import sys
import argparse
import time
import difflib
from typing import Union, List
import pandas as pd
import numpy as np


def clean_key(s: pd.Series) -> pd.Series:
    """Normalizes key strings: lowercases, strips outer spaces, collapses internal whitespace."""
    return (
        s.fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .apply(lambda x: " ".join(x.split()))
        .replace({"nan": "", "none": ""})
    )


def is_value_empty(val) -> bool:
    """Checks if a value is empty, null, NaN, or whitespace."""
    if pd.isna(val):
        return True
    s = str(val).strip()
    return not s or s.lower() in ("nan", "none", "null", r"\n", r"\\n", "[unmatched]", "unmatched")


def load_file(file_path: str, sheet_name=0) -> pd.DataFrame:
    """Loads CSV or Excel file into a DataFrame."""
    _, ext = os.path.splitext(file_path.lower())
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(file_path, sheet_name=sheet_name)
    elif ext == ".csv":
        return pd.read_csv(file_path, low_memory=False)
    else:
        raise ValueError(f"Unsupported file format '{ext}'. Expected .xlsx, .xls, or .csv.")


def save_excel_formatted(df: pd.DataFrame, file_path: str):
    """Saves DataFrame to Excel or CSV based on extension."""
    os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
    _, ext = os.path.splitext(file_path.lower())
    if ext == ".csv":
        df.to_csv(file_path, index=False)
        return
    try:
        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
            workbook = writer.book
            worksheet = writer.sheets["Sheet1"]
            worksheet.hide_gridlines(2)

            header_format = workbook.add_format({
                "bold": True,
                "font_name": "Calibri",
                "font_size": 11,
                "font_color": "#333333",
                "bg_color": "#F2F2F2",
                "align": "left",
                "valign": "vcenter",
                "bottom": 1,
                "bottom_color": "#CCCCCC",
            })
            data_format = workbook.add_format({
                "bold": False,
                "font_name": "Calibri",
                "font_size": 10,
                "font_color": "#000000",
                "align": "left",
                "valign": "vcenter",
            })

            for col_idx, col_name in enumerate(df.columns):
                worksheet.write(0, col_idx, col_name, header_format)
                max_len = max(
                    df[col_name].apply(lambda x: len(str(x)) if pd.notna(x) else 0).max() if len(df) > 0 else 0,
                    len(str(col_name)),
                )
                col_width = int(min(max(max_len + 3, 12), 50))
                worksheet.set_column(col_idx, col_idx, col_width, data_format)
    except Exception as e:
        print(f" (Formatting fallback: {e})", file=sys.stderr)
        df.to_excel(file_path, index=False)


def print_ascii_table(headers: List[str], rows: List[List[str]]):
    """Prints a list of rows as a beautiful, dense ASCII table."""
    if not rows:
        return
    widths = [len(h) for h in headers]
    for r in rows:
        for i, val in enumerate(r):
            widths[i] = max(widths[i], len(str(val)))
    border = "+" + "+".join(["-" * (w + 2) for w in widths]) + "+"
    print(border)
    print("|" + "|".join([f" {headers[i].ljust(widths[i])} " for i in range(len(headers))]) + "|")
    print(border)
    for r in rows:
        print("|" + "|".join([f" {str(r[i]).ljust(widths[i])} " for i in range(len(r))]) + "|")
    print(border)


def find_column_by_names(df: pd.DataFrame, possible_names: List[str]) -> Union[str, None]:
    """Finds a column in df matching any of the possible_names (case & whitespace insensitive)."""
    cols_map = {str(c).strip().lower(): c for c in df.columns}
    for p in possible_names:
        matched = cols_map.get(p.strip().lower())
        if matched:
            return matched
    return None


def process_product_normalization(
    policy_path: Union[str, List[str]],
    norm_path: str,
    output_dir: str = None,
    in_place: bool = False,
    policy_prod_col: str = "Product",
    norm_prod_col: str = "Product",
    norm_target_col: str = "New Product Name",
    norm_carrier_col: str = "Carrier",
) -> Union[str, List[str]]:
    start_time = time.time()

    # Load normalization reference file
    if not os.path.exists(norm_path):
        raise FileNotFoundError(f"Normalization reference file not found: '{norm_path}'")

    print(f"📖 Loading Normalization Reference File: '{norm_path}'...")
    df_norm = load_file(norm_path)

    # Auto-discover reference columns
    n_prod_matched = find_column_by_names(df_norm, [norm_prod_col, "Product", "Product Name"])
    n_target_matched = find_column_by_names(df_norm, [norm_target_col, "New Product Name", "Product", "Normalized Product Name"])
    n_carrier_matched = find_column_by_names(df_norm, [norm_carrier_col, "Carrier", "New Carrier Name", "Carrier Name"])

    if not n_prod_matched:
        raise KeyError(f"Product matching column not found in normalization file. Available: {list(df_norm.columns)}")
    if not n_target_matched:
        n_target_matched = n_prod_matched  # Fallback to product column if New Product Name is absent

    print(f"  Reference Product Key Col : '{n_prod_matched}'")
    print(f"  Reference New Product Col : '{n_target_matched}'")
    if n_carrier_matched:
        print(f"  Reference Carrier Col     : '{n_carrier_matched}'")

    # Clean reference lookup dataset
    df_norm_clean = df_norm.dropna(subset=[n_prod_matched]).copy()
    df_norm_clean["_norm_key"] = clean_key(df_norm_clean[n_prod_matched])
    df_norm_clean = df_norm_clean[df_norm_clean["_norm_key"] != ""].drop_duplicates(subset=["_norm_key"], keep="first")

    # Build reference mapping dictionaries
    prod_map = dict(zip(df_norm_clean["_norm_key"], df_norm_clean[n_target_matched]))
    carrier_map = dict(zip(df_norm_clean["_norm_key"], df_norm_clean[n_carrier_matched])) if n_carrier_matched else {}

    ref_norm_keys = list(df_norm_clean["_norm_key"].unique())
    neg_words = {'unknown', 'unassigned', 'none', 'null', 'nan', 'unmatched', 'missing'}
    fuzzy_cache = {}

    def resolve_key(tk):
        if not tk or pd.isna(tk):
            return None
        tk_str = str(tk).strip()
        if not tk_str:
            return None
        if tk_str in ref_norm_keys:
            return tk_str
        if tk_str in fuzzy_cache:
            return fuzzy_cache[tk_str]
        close = difflib.get_close_matches(tk_str, ref_norm_keys, n=1, cutoff=0.80)
        if close:
            cand = close[0]
            val_neg = any(w in tk_str for w in neg_words)
            cand_neg = any(w in cand for w in neg_words)
            res = cand if (val_neg == cand_neg) else None
        else:
            res = None
        fuzzy_cache[tk_str] = res
        return res

    # Collect target files
    target_files = []
    if isinstance(policy_path, list):
        for p in policy_path:
            if os.path.isdir(p):
                for root, _, files in os.walk(p):
                    for f in files:
                        if not f.startswith(".") and f.lower().endswith((".csv", ".xlsx", ".xls")):
                            target_files.append(os.path.join(root, f))
            elif os.path.exists(p):
                target_files.append(p)
    elif os.path.isdir(policy_path):
        for root, _, files in os.walk(policy_path):
            for f in files:
                if not f.startswith(".") and f.lower().endswith((".csv", ".xlsx", ".xls")):
                    target_files.append(os.path.join(root, f))
    elif os.path.exists(policy_path):
        target_files.append(policy_path)

    if not target_files:
        raise FileNotFoundError(f"No valid Excel/CSV policy files found for path '{policy_path}'.")

    print(f"📖 Found {len(target_files)} target policy file(s) to process.")

    # Determine destination directory
    if not in_place and not output_dir:
        if isinstance(policy_path, str) and os.path.isdir(policy_path):
            output_dir = policy_path.rstrip(os.sep) + "_normalized"
        elif len(target_files) == 1:
            output_dir = os.path.dirname(os.path.abspath(target_files[0]))
        else:
            output_dir = os.path.dirname(os.path.abspath(target_files[0]))

    if output_dir and not in_place:
        os.makedirs(output_dir, exist_ok=True)

    summary_rows = []
    output_paths = []

    for t_file in target_files:
        t_name = os.path.basename(t_file)
        try:
            df_policy = load_file(t_file)
        except Exception as e:
            print(f"  (Warning: Failed to load '{t_name}': {e}, skipping)", file=sys.stderr)
            continue

        p_prod_matched = find_column_by_names(df_policy, [policy_prod_col, "Product", "Product Name", "Product_Description"])
        p_carrier_matched = find_column_by_names(df_policy, ["Carrier", "Carrier Name"])

        if not p_prod_matched:
            print(f"  (Warning: Product column missing in '{t_name}'. Skipping)", file=sys.stderr)
            continue

        df_work = df_policy.copy()
        df_work["_raw_key"] = clean_key(df_work[p_prod_matched])
        df_work["_norm_key"] = df_work["_raw_key"].apply(resolve_key)

        mapped_product = df_work["_norm_key"].map(prod_map)
        is_target_empty = mapped_product.apply(is_value_empty)
        is_valid_product = df_work["_norm_key"].notna() & (~is_target_empty)

        # Update Product column: replace with New Product Name where valid, else keep original
        df_policy.loc[is_valid_product, p_prod_matched] = mapped_product[is_valid_product]
        product_matches = int(is_valid_product.sum())

        # Update or add Carrier column if carrier map exists
        carrier_matches = 0
        if carrier_map:
            mapped_carrier = df_work["_norm_key"].map(carrier_map)
            is_valid_carrier = df_work["_norm_key"].notna() & (~mapped_carrier.apply(is_value_empty))
            carrier_matches = int(is_valid_carrier.sum())

            if p_carrier_matched:
                df_policy.loc[is_valid_carrier, p_carrier_matched] = mapped_carrier[is_valid_carrier]
            else:
                df_policy["Carrier"] = mapped_carrier

        # Save output
        if in_place:
            out_file_path = t_file
        else:
            base_t, ext_t = os.path.splitext(t_name)
            if len(target_files) == 1 and not os.path.isdir(str(policy_path)):
                out_file_path = os.path.join(output_dir, f"{base_t}_Product_Normalized{ext_t}")
            else:
                out_file_path = os.path.join(output_dir, t_name)

        save_excel_formatted(df_policy, out_file_path)
        output_paths.append(out_file_path)

        summary_rows.append([
            t_name,
            f"{len(df_policy):,}",
            f"{product_matches:,}/{len(df_policy):,}",
            f"{carrier_matches:,}/{len(df_policy):,}" if carrier_map else "N/A",
            out_file_path
        ])

    elapsed = round(time.time() - start_time, 2)
    print("\nProduct & Carrier Normalization Summary:")
    print_ascii_table(["Target File", "Total Rows", "Product Matches", "Carrier Matches", "Output Path"], summary_rows)
    print(f"✅ Processed {len(summary_rows)} file(s) in {elapsed}s.\n")

    return output_paths[0] if len(output_paths) == 1 else output_paths


def main():
    parser = argparse.ArgumentParser(
        description="Normalize 'Product' and update/add 'Carrier' in policy upload file(s) or folder(s) against a Product Normalization reference file."
    )
    parser.add_argument(
        "-p", "--policy-file", "--policy-files", "-t", "--target",
        nargs="+",
        required=True,
        help="Path to target Policy Excel/CSV file(s) or folder(s).",
    )
    parser.add_argument(
        "-n", "--norm-file", "-s", "--source",
        type=str,
        required=True,
        help="Path to Product Normalization reference Excel or CSV file.",
    )
    parser.add_argument(
        "-o", "--output-dir",
        type=str,
        default=None,
        help="Path to save output normalized Excel/CSV files.",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Overwrite original target file(s) directly instead of saving to a new directory/file.",
    )
    parser.add_argument(
        "--policy-prod-col",
        type=str,
        default="Product",
        help="Product column name in policy upload file (default: 'Product').",
    )
    parser.add_argument(
        "--norm-prod-col",
        type=str,
        default="Product",
        help="Product column name in normalization file (default: 'Product').",
    )
    parser.add_argument(
        "--norm-target-col",
        type=str,
        default="New Product Name",
        help="New product target column name in normalization file (default: 'New Product Name').",
    )
    parser.add_argument(
        "--norm-carrier-col",
        type=str,
        default="Carrier",
        help="Carrier column name in normalization file (default: 'Carrier').",
    )

    args = parser.parse_args()

    policy_input = args.policy_file[0] if len(args.policy_file) == 1 else args.policy_file

    try:
        process_product_normalization(
            policy_path=policy_input,
            norm_path=args.norm_file,
            output_dir=args.output_dir,
            in_place=args.in_place,
            policy_prod_col=args.policy_prod_col,
            norm_prod_col=args.norm_prod_col,
            norm_target_col=args.norm_target_col,
            norm_carrier_col=args.norm_carrier_col,
        )
    except Exception as e:
        print(f"\n❌ Error during product normalization: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
