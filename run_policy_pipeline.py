#!/usr/bin/env python3
"""
Unified Policy Processing Pipeline

Orchestrates three steps:
1. Enriches the APL Base Policy List using the MLB Policy Master v1 file.
2. Adjusts Commission Splits and populates Agent Levels using the Step 1 output as lookup.
3. Splits only the Step 1 output file (Bulk Policy Import File) by Carrier.
"""

import argparse
import os
import sys
import subprocess
import time


def run_cmd(cmd_args: list):
    """Runs a shell command and prints output, exiting on failure."""
    print(f"\nCommand: {' '.join(cmd_args)}")
    result = subprocess.run(cmd_args, capture_output=False, text=True)
    if result.returncode != 0:
        print(f"\n❌ Error: Command failed with exit code {result.returncode}", file=sys.stderr)
        sys.exit(result.returncode)


def main():
    parser = argparse.ArgumentParser(
        description="Unified Pipeline for Policy Enrichment, Commission Split Calculations, and Carrier Splitting."
    )
    parser.add_argument(
        "--base-policy",
        type=str,
        required=True,
        help="Path to the base APL Policy List Excel file.",
    )
    parser.add_argument(
        "--master-policy",
        type=str,
        required=True,
        help="Path to the MLB Policy Master Excel file (v1).",
    )
    parser.add_argument(
        "--base-split",
        type=str,
        required=True,
        help="Path to the base APL Commission Split CSV file.",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        required=True,
        help="Path to the directory where all pipeline outputs will be saved.",
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
        help="Skip the splitting step (Step 3).",
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
        help="Path to optional reference/mapping file (Excel/CSV) to replace/update column values.",
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
        help="Value column in reference file to pull values from (default: auto-discovers matching columns).",
    )
    parser.add_argument(
        "--ref-target-col",
        "--target-col",
        "--agent-roster-target-col",
        "--roster-target-col",
        dest="ref_target_col",
        type=str,
        default=None,
        help="Target column in output dataset to update (default: same as ref-val-col).",
    )
    parser.add_argument(
        "--export-errors",
        action="store_true",
        help="Separate records with missing carrier or unmatched agent ID into a separate error Excel file.",
    )
    parser.add_argument(
        "--error-file",
        type=str,
        default=None,
        help="Path for the error Excel file when --export-errors is used.",
    )
    parser.add_argument(
        "--yes",
        "-y",
        "--auto-confirm",
        action="store_true",
        help="Auto-confirm prompt and proceed to next steps without stopping.",
    )

    args = parser.parse_args()

    # Create output directory if it doesn't exist
    if not os.path.exists(args.output_dir):
        os.makedirs(args.output_dir, exist_ok=True)

    # Derive individual output paths
    step1_out = os.path.join(args.output_dir, "Policy_List_Enriched.xlsx")
    comm_import_dir = os.path.join(args.output_dir, "commission_import")
    split_dir = os.path.join(args.output_dir, "split_policy")

    start_time = time.time()
    print("==========================================================================")
    print("🌟 STARTING UNIFIED POLICY PROCESSING PIPELINE 🌟")
    print("==========================================================================")

    # --- Step 1: Bulk Policy List Enrichment & Validation ---
    print("\n--- [STEP 1/3] Bulk Policy List Enrichment & Validation ---")
    step1_script = "update_policy_import_with_validation.py" if (args.export_errors or args.error_file) else "update_policy_import.py"
    step1_args = [
        sys.executable,
        step1_script,
        "--base-file", args.base_policy,
        "--master-file", args.master_policy,
        "--output-file", step1_out
    ]
    if args.error_file:
        step1_args.extend(["--error-file", args.error_file])
    if args.ref_file:
        step1_args.extend(["--ref-file", args.ref_file])
        if args.ref_key_col:
            step1_args.extend(["--ref-key-col", args.ref_key_col])
        if args.ref_val_col:
            step1_args.extend(["--ref-val-col", args.ref_val_col])
        if args.ref_target_col:
            step1_args.extend(["--ref-target-col", args.ref_target_col])
    run_cmd(step1_args)

    # Check if errors were detected in Step 1
    error_count = 0
    eff_error_file = args.error_file or os.path.join(args.output_dir, "Policy_List_Enriched_Errors.xlsx")
    if os.path.exists(eff_error_file):
        try:
            import pandas as pd
            err_df_check = pd.read_excel(eff_error_file)
            error_count = len(err_df_check)
        except Exception:
            error_count = 0

    if error_count == 0:
        print("\n✨ Zero errors found in Step 1! Automatically proceeding to next pipeline steps without interruption...\n")
    else:
        # Interactive prompt to continue or break when errors exist
        if not args.yes and sys.stdin.isatty():
            try:
                resp = input(f"\n❓ Errors were detected in Step 1 ({error_count:,} record(s)). Do you want to proceed to Commission Split Calculation & Carrier Splitting? [y/N]: ").strip().lower()
                if resp not in ("y", "yes"):
                    print("\n🛑 Pipeline paused by user. Stopping after Step 1 Policy List Enrichment.\n")
                    sys.exit(0)
            except (EOFError, KeyboardInterrupt):
                print("\n")
                sys.exit(0)

    # --- Step 2: Commission Split Calculation & Agent Level Enrichment ---
    print("\n--- [STEP 2/3] Commission Split Calculation & Agent Level Enrichment ---")
    step2_args = [
        sys.executable,
        "enrich_policy_commission_import.py",
        "--base-file", args.base_split,
        "--master-file", step1_out,  # Use step 1 output as lookup!
        "--output-dir", comm_import_dir
    ]
    if not os.path.isdir(args.base_split):
        base_name = os.path.splitext(os.path.basename(args.base_split))[0]
        step2_out = os.path.join(comm_import_dir, f"{base_name}_enriched.csv")
        step2_args.extend(["--output-file", step2_out])

    if args.ref_file:
        step2_args.extend(["--ref-file", args.ref_file])
        if args.ref_key_col:
            step2_args.extend(["--ref-key-col", args.ref_key_col])
        if args.ref_val_col:
            step2_args.extend(["--ref-val-col", args.ref_val_col])
        if args.ref_target_col:
            step2_args.extend(["--ref-target-col", args.ref_target_col])
    run_cmd(step2_args)

    # --- Step 3: Split Policy List (Step 1 Output) by Carrier ---
    if not args.no_splits:
        print("\n--- [STEP 3/3] Carrier Splitting (Policy Import File) ---")
        step3_args = [
            sys.executable,
            "split_file.py",
            "--file", step1_out,
            "--column", args.split_col,
            "--output-dir", split_dir,
            "--format", "match"
        ]
        run_cmd(step3_args)
    else:
        print("\n--- [STEP 3/3] Carrier Splitting: Skipped ---")

    elapsed_time = time.time() - start_time
    print("\n==========================================================================")
    print("✅ PIPELINE EXECUTION SUCCESSFUL")
    print(f" Total Time: {elapsed_time:.1f} seconds")
    print("==========================================================================")


if __name__ == "__main__":
    main()
