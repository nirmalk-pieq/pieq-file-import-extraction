#!/usr/bin/env python3
"""Run APL vs ICM payables comparison into an explicit output folder (no versions/)."""

from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path


def load_validator():
    script = (
        Path(__file__).resolve().parents[2]
        / "apl_icm_cross_validation"
        / "validate_apl_icm.py"
    )
    if not script.is_file():
        raise FileNotFoundError(f"validate_apl_icm.py not found at {script}")
    spec = importlib.util.spec_from_file_location("validate_apl_icm", script)
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not load {script}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Compare APL statement payables to ICM / Parallel payables."
    )
    parser.add_argument("--apl", required=True, help="APL statement CSV or Excel path")
    parser.add_argument("--parallel", required=True, help="ICM / Parallel payables CSV or Excel path")
    parser.add_argument("--out-dir", required=True, help="Output folder for this comparison run")
    parser.add_argument("--mode", choices=("apl", "icm"), default="icm")
    parser.add_argument("--tolerance", type=float, default=0.02)
    args = parser.parse_args()

    apl = Path(args.apl).expanduser().resolve()
    parallel = Path(args.parallel).expanduser().resolve()
    out_dir = Path(args.out_dir).expanduser().resolve()

    if not apl.is_file():
        print(f"Error: APL file not found: {apl}", file=sys.stderr)
        return 1
    if not parallel.is_file():
        print(f"Error: ICM / Parallel file not found: {parallel}", file=sys.stderr)
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    out_csv = out_dir / "Validation_Output.csv"
    out_xlsx = out_dir / "Validation_Output.xlsx"

    print(f"Payables comparison → {out_dir}")
    print(f"APL: {apl}")
    print(f"ICM: {parallel}")
    print(f"Mode: {args.mode}  Tolerance: {args.tolerance}")

    validator = load_validator()
    validator.validate_data(
        str(apl),
        str(parallel),
        str(out_csv),
        str(out_xlsx),
        amount_tolerance=args.tolerance,
        mode=args.mode,
        single_grouped_file=True,
    )
    print(f"Outputs written under {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
