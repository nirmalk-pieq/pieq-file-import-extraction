#!/usr/bin/env python3
"""
Automated Policy Product Normalization & Carrier Mapping Pipeline

This script automates:
1. Product name normalization using a master reference dictionary.
2. Asterisk (*, **), symbol, and punctuation discrepancy resolution.
3. Carrier column updates based on the product-carrier catalog.

Supports both Excel (.xlsx) and CSV (.csv) files.
"""

import argparse
import csv
import glob
import os
import re
import sys
import time
from pathlib import Path
import openpyxl

DEFAULT_RAW_DIR = "/Users/nirmalrajaak/pieq/Insurance/pieq-policy-file-import-extraction/version_3/(AUG 3)RAW MASTER DATA for POLICY COMMISSIONS IMPORT 2"
DEFAULT_NORM_FILE = os.path.join(DEFAULT_RAW_DIR, "Normalised Product Name-(728 Products ,unique 57K policies).xlsx")
DEFAULT_CARRIER_FILE = os.path.join(DEFAULT_RAW_DIR, "Product-Carrier Sheet.xlsx")


def norm_space(s: str) -> str:
    """Collapses multiple whitespace characters to a single space and strips ends."""
    if s is None:
        return ""
    return re.sub(r'\s+', ' ', str(s)).strip()


def strip_suffixes(s: str) -> str:
    """Strips trailing asterisks, exclamations, hyphens, and whitespace."""
    if s is None:
        return ""
    return re.sub(r'[\*\!\-\+\#\s]+$', '', str(s)).strip()


def clean_core(s: str) -> str:
    """Replaces non-alphanumeric characters with spaces and collapses."""
    if s is None:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', ' ', str(s)).strip()


def pd_is_na(val):
    if val is None:
        return True
    s = str(val).strip().lower()
    return s in ('nan', 'none', 'null', '')


class ProductNormalizer:
    """6-Tier Smart Product Matching Engine"""

    def __init__(self, norm_excel_path: str):
        self.norm_excel_path = norm_excel_path
        self.exact_map = {}
        self.strip_map = {}
        self.space_map = {}
        self.lower_map = {}
        self.symbol_map = {}
        self.core_map = {}
        self._load_master_reference()

    def _load_master_reference(self):
        wb = openpyxl.load_workbook(self.norm_excel_path, read_only=True)
        sheet = wb.active
        headers = [str(cell.value).strip() if cell.value else "" for cell in next(sheet.iter_rows(max_row=1))]
        
        prod_idx = -1
        new_prod_idx = -1
        for i, h in enumerate(headers):
            if h.lower() == 'product':
                prod_idx = i
            elif h.lower() == 'new product name':
                new_prod_idx = i

        if prod_idx == -1 or new_prod_idx == -1:
            raise ValueError(f"Could not find required columns ('Product', 'New Product Name') in {self.norm_excel_path}")

        count = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or len(row) <= max(prod_idx, new_prod_idx):
                continue
            p_val = row[prod_idx]
            np_val = row[new_prod_idx]

            if p_val is None or pd_is_na(p_val):
                continue
            p_raw = str(p_val)
            np_str = str(np_val).strip() if np_val is not None else ""

            if not p_raw or p_raw == 'nan' or not np_str:
                continue

            count += 1
            self.exact_map[p_raw] = np_str
            p_trim = p_raw.strip()
            self.strip_map[p_trim] = np_str
            p_space = norm_space(p_raw)
            self.space_map[p_space] = np_str
            self.lower_map[p_space.lower()] = np_str

            p_sym = strip_suffixes(p_raw).lower()
            if p_sym and p_sym not in self.symbol_map:
                self.symbol_map[p_sym] = np_str

            p_core = clean_core(p_raw).lower()
            if p_core and p_core not in self.core_map:
                self.core_map[p_core] = np_str

        wb.close()
        print(f"  [ProductNormalizer] Indexed {count} reference product entries.")

    def resolve(self, val) -> str:
        """Resolves raw product name using 6-tier matching engine."""
        if val is None or pd_is_na(val):
            return None
        val_str = str(val)
        if not val_str or val_str == 'nan':
            return None

        if val_str in self.exact_map:
            return self.exact_map[val_str]
        val_trim = val_str.strip()
        if val_trim in self.strip_map:
            return self.strip_map[val_trim]
        val_space = norm_space(val_str)
        if val_space in self.space_map:
            return self.space_map[val_space]
        if val_space.lower() in self.lower_map:
            return self.lower_map[val_space.lower()]
        val_sym = strip_suffixes(val_str).lower()
        if val_sym in self.symbol_map:
            return self.symbol_map[val_sym]
        val_core = clean_core(val_str).lower()
        if val_core in self.core_map:
            return self.core_map[val_core]

        return None


class CarrierMapper:
    """Catalog-Based Product to Carrier Mapping Engine"""

    def __init__(self, carrier_excel_path: str, normalizer: ProductNormalizer):
        self.carrier_excel_path = carrier_excel_path
        self.normalizer = normalizer
        self.carrier_map = {}
        self._load_carrier_catalog()

    def _load_carrier_catalog(self):
        wb = openpyxl.load_workbook(self.carrier_excel_path, read_only=True)
        sheet = wb.active
        headers = [str(cell.value).strip() if cell.value else "" for cell in next(sheet.iter_rows(max_row=1))]

        prod_idx = -1
        carrier_idx = -1
        for i, h in enumerate(headers):
            if h.lower() == 'product':
                prod_idx = i
            elif 'carrier' in h.lower():
                carrier_idx = i

        if prod_idx == -1 or carrier_idx == -1:
            raise ValueError(f"Could not find required columns ('Product', Carrier) in {self.carrier_excel_path}")

        count = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or len(row) <= max(prod_idx, carrier_idx):
                continue
            p_val = row[prod_idx]
            c_val = row[carrier_idx]

            if p_val is None or c_val is None or pd_is_na(p_val) or pd_is_na(c_val):
                continue

            p_raw = str(p_val)
            c_str = str(c_val).strip()
            if not p_raw or p_raw == 'nan' or not c_str or c_str == 'nan':
                continue

            count += 1
            self.carrier_map[p_raw] = c_str
            self.carrier_map[p_raw.strip()] = c_str
            p_space = norm_space(p_raw)
            self.carrier_map[p_space] = c_str
            self.carrier_map[p_space.lower()] = c_str

        wb.close()
        print(f"  [CarrierMapper] Indexed {count} catalog carrier entries.")

    def resolve_carrier(self, product_val) -> str:
        if product_val is None or pd_is_na(product_val):
            return None
        val_str = str(product_val)
        if val_str in self.carrier_map:
            return self.carrier_map[val_str]
        val_strip = val_str.strip()
        if val_strip in self.carrier_map:
            return self.carrier_map[val_strip]
        val_space = norm_space(val_str)
        if val_space in self.carrier_map:
            return self.carrier_map[val_space]
        if val_space.lower() in self.carrier_map:
            return self.carrier_map[val_space.lower()]

        # Try resolving normalized product first, then mapping carrier
        norm_p = self.normalizer.resolve(product_val)
        if norm_p and norm_p in self.carrier_map:
            return self.carrier_map[norm_p]
        if norm_p and norm_p.lower() in self.carrier_map:
            return self.carrier_map[norm_p.lower()]

        return None


class FileProcessor:
    """Processes target Excel and CSV policy files"""

    def __init__(self, normalizer: ProductNormalizer, mapper: CarrierMapper, dry_run: bool = False):
        self.normalizer = normalizer
        self.mapper = mapper
        self.dry_run = dry_run

    def process_file(self, file_path: str) -> dict:
        ext = Path(file_path).suffix.lower()
        if ext == '.xlsx':
            return self._process_excel(file_path)
        elif ext == '.csv':
            return self._process_csv(file_path)
        else:
            print(f"Skipping unsupported file format: {file_path}")
            return {'scanned': 0, 'products_fixed': 0, 'carriers_fixed': 0}

    def _process_excel(self, file_path: str) -> dict:
        filename = os.path.basename(file_path)
        print(f"\nProcessing Excel: {filename}...")
        t0 = time.time()
        
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        headers = [str(cell.value).strip() if cell.value else "" for cell in sheet[1]]

        prod_col_idx = -1
        carrier_col_idx = -1

        for i, h in enumerate(headers):
            if h.lower() == 'product':
                prod_col_idx = i + 1
            elif h.lower() == 'carrier':
                carrier_col_idx = i + 1

        if prod_col_idx == -1:
            print(f"  [Warning] No 'Product' column in {filename}. Skipping.")
            wb.close()
            return {'scanned': 0, 'products_fixed': 0, 'carriers_fixed': 0}

        scanned = sheet.max_row - 1
        prods_fixed = 0
        carriers_fixed = 0

        for row_idx in range(2, sheet.max_row + 1):
            prod_cell = sheet.cell(row=row_idx, column=prod_col_idx)
            if prod_cell.value is not None:
                new_p = self.normalizer.resolve(prod_cell.value)
                if new_p is not None and prod_cell.value != new_p:
                    prod_cell.value = new_p
                    prods_fixed += 1

            if carrier_col_idx != -1:
                carrier_cell = sheet.cell(row=row_idx, column=carrier_col_idx)
                target_p = prod_cell.value
                new_c = self.mapper.resolve_carrier(target_p)
                if new_c is not None and str(carrier_cell.value).strip() != new_c:
                    carrier_cell.value = new_c
                    carriers_fixed += 1

        if not self.dry_run:
            wb.save(file_path)
            print(f"  Saved {filename} in {time.time() - t0:.2f}s (Products updated: {prods_fixed}, Carriers updated: {carriers_fixed})")
        else:
            wb.close()
            print(f"  [Dry-Run] {filename} analyzed in {time.time() - t0:.2f}s (Products to update: {prods_fixed}, Carriers to update: {carriers_fixed})")

        return {'scanned': scanned, 'products_fixed': prods_fixed, 'carriers_fixed': carriers_fixed}

    def _process_csv(self, file_path: str) -> dict:
        filename = os.path.basename(file_path)
        print(f"\nProcessing CSV: {filename}...")
        t0 = time.time()

        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            reader = list(csv.reader(f))

        if not reader:
            return {'scanned': 0, 'products_fixed': 0, 'carriers_fixed': 0}

        headers = [str(h).strip().replace('"', '') for h in reader[0]]
        prod_idx = -1
        carrier_idx = -1

        for i, h in enumerate(headers):
            if h.lower() == 'product':
                prod_idx = i
            elif h.lower() == 'carrier':
                carrier_idx = i

        if prod_idx == -1:
            print(f"  [Warning] No 'Product' column in {filename}. Skipping.")
            return {'scanned': 0, 'products_fixed': 0, 'carriers_fixed': 0}

        scanned = len(reader) - 1
        prods_fixed = 0
        carriers_fixed = 0

        for row in reader[1:]:
            if len(row) <= prod_idx:
                continue
            cur_p = row[prod_idx]
            new_p = self.normalizer.resolve(cur_p)
            if new_p is not None and cur_p != new_p:
                row[prod_idx] = new_p
                prods_fixed += 1

            if carrier_idx != -1 and len(row) > carrier_idx:
                cur_c = row[carrier_idx]
                target_p = row[prod_idx]
                new_c = self.mapper.resolve_carrier(target_p)
                if new_c is not None and cur_c.strip() != new_c:
                    row[carrier_idx] = new_c
                    carriers_fixed += 1

        if not self.dry_run:
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerows(reader)
            print(f"  Saved {filename} in {time.time() - t0:.2f}s (Products updated: {prods_fixed}, Carriers updated: {carriers_fixed})")
        else:
            print(f"  [Dry-Run] {filename} analyzed in {time.time() - t0:.2f}s (Products to update: {prods_fixed}, Carriers to update: {carriers_fixed})")

        return {'scanned': scanned, 'products_fixed': prods_fixed, 'carriers_fixed': carriers_fixed}


def main():
    parser = argparse.ArgumentParser(description="Policy Product Normalization & Carrier Mapping Pipeline")
    parser.add_argument("--norm-file", default=DEFAULT_NORM_FILE, help="Path to normalized product reference Excel file")
    parser.add_argument("--carrier-file", default=DEFAULT_CARRIER_FILE, help="Path to product-carrier catalog Excel file")
    parser.add_argument("--target-file", help="Path to a single target .xlsx or .csv file")
    parser.add_argument("--target-dir", help="Directory containing target files to process")
    parser.add_argument("--file-pattern", default="*.xlsx", help="Glob pattern for target files (default: *.xlsx)")
    parser.add_argument("--dry-run", action="store_true", help="Scan and report without modifying files on disk")

    args = parser.parse_args()

    print("============================================================")
    print("POLICY PRODUCT & CARRIER AUTO-NORMALIZATION PIPELINE")
    print("============================================================")
    if args.dry_run:
        print("[MODE]: DRY-RUN (No files will be modified on disk)\n")

    t_start = time.time()

    normalizer = ProductNormalizer(args.norm_file)
    mapper = CarrierMapper(args.carrier_file, normalizer)
    processor = FileProcessor(normalizer, mapper, dry_run=args.dry_run)

    target_files = []
    if args.target_file:
        if os.path.exists(args.target_file):
            target_files.append(args.target_file)
        else:
            print(f"Error: Target file not found: {args.target_file}")
            sys.exit(1)
    elif args.target_dir:
        search_path = os.path.join(args.target_dir, args.file_pattern)
        target_files = glob.glob(search_path)
    else:
        # Default target files in standard raw folder
        default_targets = ["MLB_policy_master.xlsx", "APL_Policy_List.xlsx", "Product-Carrier Sheet.xlsx"]
        for t in default_targets:
            full_p = os.path.join(DEFAULT_RAW_DIR, t)
            if os.path.exists(full_p):
                target_files.append(full_p)

    if not target_files:
        print("No target files found to process.")
        sys.exit(0)

    total_scanned = 0
    total_prods_fixed = 0
    total_carriers_fixed = 0

    for tf in target_files:
        stats = processor.process_file(tf)
        total_scanned += stats['scanned']
        total_prods_fixed += stats['products_fixed']
        total_carriers_fixed += stats['carriers_fixed']

    t_elapsed = time.time() - t_start
    print("\n============================================================")
    print("PIPELINE SUMMARY REPORT")
    print("============================================================")
    print(f"Total Files Processed : {len(target_files)}")
    print(f"Total Rows Scanned    : {total_scanned:,}")
    print(f"Total Products Fixed  : {total_prods_fixed:,}")
    print(f"Total Carriers Fixed  : {total_carriers_fixed:,}")
    print(f"Pipeline Status       : SUCCESS (Completed in {t_elapsed:.2f}s)")
    print("============================================================")


if __name__ == '__main__':
    main()
