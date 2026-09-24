#!/usr/bin/env python3
import os
import sys
import pandas as pd

def main():
    file_a = "error_checking.xlsx"
    file_b = "(AUG 3)PROCESSED MASTER DATA for POLICY COMMISSIONS IMPORT/v2_last_final_commission_import/v4_validated/commission_corrected_clean.xlsx"
    output_dir = "(AUG 3)PROCESSED MASTER DATA for POLICY COMMISSIONS IMPORT/v2_last_final_commission_import/v4_validated/original"
    output_file = os.path.join(output_dir, "commission_corrected_clean.xlsx")

    if not os.path.exists(file_a):
        print(f"❌ Error: File A not found: {file_a}", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(file_b):
        print(f"❌ Error: File B not found: {file_b}", file=sys.stderr)
        sys.exit(1)

    print(f"Reading File A (error_checking): {file_a}...")
    df_a = pd.read_excel(file_a)
    print(f"Reading File B (commission_corrected_clean): {file_b}...")
    df_b = pd.read_excel(file_b)

    # Standardize columns to string and strip spaces
    for col in ["Policy No", "Carrier"]:
        df_a[col] = df_a[col].astype(str).str.strip()
        df_b[col] = df_b[col].astype(str).str.strip()

    # Drop duplicates in File A keys
    df_a_keys = df_a[["Policy No", "Carrier"]].drop_duplicates()

    # To remove rows from B that match keys in A:
    # We do a left join with an indicator, then filter where indicator is 'left_only'
    merged = pd.merge(df_b, df_a_keys, on=["Policy No", "Carrier"], how="left", indicator=True)
    unmatched_df = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"])

    print(f"🎉 Filtering completed!")
    print(f"- Total rows in original File B: {len(df_b)}")
    print(f"- Rows removed (matched): {len(df_b) - len(unmatched_df)}")
    print(f"- Remaining rows (unmatched): {len(unmatched_df)}")

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    try:
        unmatched_df.to_excel(output_file, index=False)
        print(f"✅ Filtered file successfully saved to: '{output_file}'")
    except Exception as e:
        print(f"❌ Error writing output file: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
