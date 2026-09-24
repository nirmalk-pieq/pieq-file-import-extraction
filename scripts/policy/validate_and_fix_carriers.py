"""
carrier_validation_and_fix.py
---------------------------------
Logic:
  1. Load commission file (all_policies_enriched_merged.csv) - source of truth for commission rows
  2. Load master file (Merged_Import.xlsx) - authoritative carrier & policy list

  3. For each commission row:
     a) If Policy No + Product NOT found in master  → write to  commission_not_in_master.csv
     b) If Policy No + Product found in master:
        - Compare commission Carrier vs master Carrier
        - If DIFFERENT → override commission Carrier with master Carrier
          (record is kept in the corrected output)
        - If SAME → keep as-is

Outputs (NEW files, originals untouched):
  - commission_corrected.csv         : all rows that matched master (carrier corrected where needed)
  - commission_carrier_changed.csv   : subset of above where carrier was actually changed (audit trail)
  - commission_not_in_master.csv     : rows whose policy number + product was not found in master
"""

import pandas as pd
import os
import argparse
import sys

def main():
    BASE_DIR = "/Users/nirmalrajaak/pieq/Insurance/pieq-policy-file-import-extraction/(AUG 3)PROCESSED MASTER DATA for POLICY COMMISSIONS IMPORT"

    default_commission = os.path.join(BASE_DIR, "commission_import", "all_policies_enriched_merged.csv")
    default_master = os.path.join(BASE_DIR, "Merged_Import.xlsx")
    default_out_dir = os.path.join(BASE_DIR, "commission_import", "validated_output")

    parser = argparse.ArgumentParser(
        description="Reconcile and correct carrier names in commission rows by comparing against a master policy file."
    )
    parser.add_argument(
        "-c", "--commission-file",
        default=default_commission,
        help=f"Path to commission CSV file (default: {default_commission})"
    )
    parser.add_argument(
        "-m", "--master-file",
        default=default_master,
        help=f"Path to master Excel/CSV file (default: {default_master})"
    )
    parser.add_argument(
        "-o", "--output-dir",
        default=default_out_dir,
        help=f"Output directory to save corrected files (default: {default_out_dir})"
    )

    args = parser.parse_args()

    commission_file = args.commission_file
    master_file = args.master_file
    out_dir = args.output_dir

    if not os.path.exists(commission_file):
        print(f"ERROR: Commission file not found: {commission_file}", file=sys.stderr)
        sys.exit(1)

    if not os.path.exists(master_file):
        print(f"ERROR: Master file not found: {master_file}", file=sys.stderr)
        sys.exit(1)

    os.makedirs(out_dir, exist_ok=True)

    out_corrected       = os.path.join(out_dir, "commission_corrected.csv")
    out_carrier_changed = os.path.join(out_dir, "commission_carrier_changed.csv")
    out_not_in_master   = os.path.join(out_dir, "commission_not_in_master.csv")

    # ── 1. Load files ──────────────────────────────────────────────────────────────
    print(f"Loading commission file: {commission_file}...")
    comm_df = pd.read_csv(commission_file, dtype=str)
    print(f"  Commission rows : {len(comm_df):,}")

    print(f"Loading master file: {master_file}...")
    _, ext = os.path.splitext(master_file.lower())
    if ext in ('.xlsx', '.xls'):
        master_df = pd.read_excel(master_file, dtype=str)
    else:
        master_df = pd.read_csv(master_file, dtype=str)
    print(f"  Master rows     : {len(master_df):,}")

    # ── 2. Build master lookup: (Policy No, Product) → Carrier (keep first occurrence) ──
    # Normalize matching fields (fill empty, strip whitespace, uppercase)
    comm_df["Policy No"] = comm_df["Policy No"].fillna("").astype(str).str.strip()
    comm_df["Product"]   = comm_df["Product"].fillna("").astype(str).str.strip().str.upper()

    master_df["Policy No"] = master_df["Policy No"].fillna("").astype(str).str.strip()
    master_df["Product"]   = master_df["Product"].fillna("").astype(str).str.strip().str.upper()
    master_df["Carrier"]   = master_df["Carrier"].fillna("").astype(str).str.strip()

    # Create normalized composite key
    comm_df["_key"]   = comm_df["Policy No"] + "||" + comm_df["Product"]
    master_df["_key"] = master_df["Policy No"] + "||" + master_df["Product"]

    # De-duplicate master on composite key – take first occurrence
    master_lookup = (
        master_df[["_key", "Carrier"]]
        .drop_duplicates(subset=["_key"], keep="first")
        .set_index("_key")["Carrier"]
    )
    print(f"\nUnique policy-product keys in master: {len(master_lookup):,}")

    # ── 3. Split commission rows into matched / not-matched ───────────────────────
    comm_df["_in_master"] = comm_df["_key"].isin(master_lookup.index)

    not_in_master = comm_df[~comm_df["_in_master"]].copy()
    matched       = comm_df[ comm_df["_in_master"]].copy()

    print(f"\nMatched rows (policy-product found in master)     : {len(matched):,}")
    print(f"Not-in-master rows (policy-product NOT in master) : {len(not_in_master):,}")

    # ── 4. For matched rows – compare & correct carrier ───────────────────────────
    matched["_master_carrier"] = matched["_key"].map(master_lookup)

    # Strip for fair comparison
    matched["Carrier"]          = matched["Carrier"].fillna("").astype(str).str.strip()
    matched["_master_carrier"]  = matched["_master_carrier"].fillna("").astype(str).str.strip()

    carrier_changed = matched["Carrier"] != matched["_master_carrier"]
    print(f"\nRows with carrier mismatch (will be corrected): {carrier_changed.sum():,}")
    print(f"Rows with carrier already matching             : {(~carrier_changed).sum():,}")

    # Audit file: rows where carrier was changed
    audit_df = matched[carrier_changed].copy()
    audit_df = audit_df.rename(columns={"Carrier": "Old_Carrier"})
    audit_df["Carrier"] = audit_df["_master_carrier"]
    # Reorder so old/new are visible side-by-side
    cols = list(comm_df.columns)
    cols.remove("_in_master")
    cols.remove("_key")
    audit_cols = []
    for c in cols:
        if c == "Carrier":
            audit_cols.append("Old_Carrier")
            audit_cols.append("Carrier")
        else:
            audit_cols.append(c)
    audit_df = audit_df[audit_cols]

    # Apply correction to all matched rows
    matched.loc[carrier_changed, "Carrier"] = matched.loc[carrier_changed, "_master_carrier"]

    # ── 5. Clean up temp columns ──────────────────────────────────────────────────
    matched       = matched.drop(columns=["_in_master", "_master_carrier", "_key"])
    not_in_master = not_in_master.drop(columns=["_in_master", "_key"])

    # ── 6. Write outputs ──────────────────────────────────────────────────────────
    print(f"\nWriting commission_corrected.csv       ({len(matched):,} rows) to: {out_corrected}")
    matched.to_csv(out_corrected, index=False)

    print(f"Writing commission_carrier_changed.csv ({len(audit_df):,} rows) to: {out_carrier_changed}")
    audit_df.to_csv(out_carrier_changed, index=False)

    print(f"Writing commission_not_in_master.csv   ({len(not_in_master):,} rows) to: {out_not_in_master}")
    not_in_master.to_csv(out_not_in_master, index=False)

    # ── 7. Summary ────────────────────────────────────────────────────────────────
    print(f"""
╔══════════════════════════════════════════════════════════════╗
║                        SUMMARY                               ║
╠══════════════════════════════════════════════════════════════╣
║  Total commission rows           : {len(comm_df):>10,}               ║
║  ├─ Matched in master            : {len(matched):>10,}               ║
║  │   ├─ Carrier already correct  : {(~carrier_changed).sum():>10,}               ║
║  │   └─ Carrier corrected        : {carrier_changed.sum():>10,}               ║
║  └─ NOT in master                : {len(not_in_master):>10,}               ║
╠══════════════════════════════════════════════════════════════╣
║  Output directory:                                           ║
║  {out_dir[:58]}  ║
╠══════════════════════════════════════════════════════════════╣
║  commission_corrected.csv        (all matched rows)          ║
║  commission_carrier_changed.csv  (audit: changed carriers)   ║
║  commission_not_in_master.csv    (policies not in master)    ║
╚══════════════════════════════════════════════════════════════╝
""")

if __name__ == "__main__":
    main()
