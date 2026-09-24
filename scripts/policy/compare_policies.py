#!/usr/bin/env python3
import argparse
import os
import sys
import pandas as pd

def main():
    parser = argparse.ArgumentParser(
        description="Compare two policy files on 'Policy No' and 'Carrier' and output the matching records."
    )
    parser.add_argument(
        "-a", "--file-a",
        default="error_checking.xlsx",
        help="Path to the first file (e.g. error_checking.xlsx)."
    )
    parser.add_argument(
        "-b", "--file-b",
        default="(AUG 3)PROCESSED MASTER DATA for POLICY COMMISSIONS IMPORT/v2_last_final_commission_import/validated/commission_corrected_clean.xlsx",
        help="Path to the second file (e.g. commission_corrected_clean.xlsx)."
    )
    parser.add_argument(
        "-o", "--output-excel",
        default="matched_policies.xlsx",
        help="Path to the output Excel file for matched records."
    )
    args = parser.parse_args()

    if not os.path.exists(args.file_a):
        print(f"❌ Error: File A not found: {args.file_a}", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(args.file_b):
        print(f"❌ Error: File B not found: {args.file_b}", file=sys.stderr)
        sys.exit(1)

    print(f"Reading File A: {args.file_a}...")
    df_a = pd.read_excel(args.file_a)
    print(f"Reading File B: {args.file_b}...")
    df_b = pd.read_excel(args.file_b)

    # Standardize columns to string and strip spaces
    for col in ["Policy No", "Carrier"]:
        if col in df_a.columns:
            df_a[col] = df_a[col].astype(str).str.strip()
        else:
            print(f"❌ Error: Column '{col}' missing in File A.", file=sys.stderr)
            sys.exit(1)

        if col in df_b.columns:
            df_b[col] = df_b[col].astype(str).str.strip()
        else:
            print(f"❌ Error: Column '{col}' missing in File B.", file=sys.stderr)
            sys.exit(1)

    # Find all matching rows in df_b (retains all sequences)
    matched_details = pd.merge(df_b, df_a[["Policy No", "Carrier"]].drop_duplicates(), on=["Policy No", "Carrier"], how="inner")

    # Generate summary sheet of policy-wise row counts
    matched_summary = matched_details.groupby(["Policy No", "Carrier"]).size().reset_index(name="Row Count")
    # Sort by Row Count descending
    matched_summary = matched_summary.sort_values(by="Row Count", ascending=False)

    print(f"🎉 Matching completed!")
    print(f"- File A unique records: {len(df_a[['Policy No', 'Carrier']].drop_duplicates())}")
    print(f"- File B total sequence rows: {len(df_b)}")
    print(f"- Matching unique Policies: {len(matched_summary)}")
    print(f"- Matching total sequence rows: {len(matched_details)}")

    try:
        with pd.ExcelWriter(args.output_excel, engine="openpyxl") as writer:
            matched_summary.to_excel(writer, sheet_name="Policy Row Counts", index=False)
            matched_details.to_excel(writer, sheet_name="Matched Sequences", index=False)
        print(f"✅ Matched list and row counts saved to: '{args.output_excel}'")
    except Exception as e:
        print(f"❌ Error saving output: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
