import pandas as pd
import numpy as np
from enum import Enum
import warnings
import os
import shutil
warnings.filterwarnings('ignore')

class APLColumns(str, Enum):
    POLICY_NUMBER = "Policy Number"
    AGENT_NAME = "Agent Name"
    PRODUCT = "Product"

AMOUNT_COLUMNS = [
    "Earned Commission",
    "Gross Commission",
    "Advanced",
    "Earned Credit",
    "Chargeback",
    "Escrow Adjustment",
    "Escrow Amount"
]

class ParallelColumns(str, Enum):
    POLICY_NUMBER = "policy_number"
    AGENT_NAME = "Agent Name"
    AMOUNT = "commission_amount"
    PRODUCT = "product_name"

import re

def normalize_name(name):
    if pd.isna(name): return set()
    name_str = re.sub(r'[,.]', '', str(name)).lower().strip()
    
    # Alias Main Line Benefits / MLB
    if 'main line benefits' in name_str or 'mlb' in name_str:
        return {'mlb'}
        
    words = name_str.split()
    if not words: return set()
    if len(words) == 1: return {words[0]}
    return {words[0], words[-1]}

def clean_amount(val):
    if pd.isna(val) or str(val).strip() == 'nan':
        return 0.0
    try:
        cleaned = str(val).replace(',', '').replace('$', '').strip()
        if not cleaned:
            return 0.0
        return float(cleaned)
    except ValueError:
        return 0.0

def check_amount_match(row, icm_amt_val, tolerance=0.01):
    for col in AMOUNT_COLUMNS:
        if col not in row: continue
        val_str = str(row[col]).strip() if pd.notnull(row[col]) else ""
        if val_str.lower() == 'nan': val_str = ""
        val_num = clean_amount(val_str)
        if abs(val_num - icm_amt_val) <= tolerance:
            return col, val_str
            
    # Default fallback
    default_str = str(row.get("Earned Commission", "")).strip()
    if default_str.lower() == 'nan': default_str = ""
    return "None", default_str

def load_file(file_path):
    if file_path.lower().endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path, dtype=str)
    else:
        try:
            return pd.read_csv(file_path, dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            return pd.read_csv(file_path, dtype=str, encoding='latin-1')

def validate_data(apl_path, parallel_path, out_csv_path, out_xlsx_path, amount_tolerance=0.01):
    print("Loading APL Statement...")
    try:
        df_apl = load_file(apl_path)
        print(f" -> Loaded {len(df_apl)} records from APL.")
    except Exception as e:
        print(f"Error loading APL file: {e}")
        return

    print("Loading Parallel Run Payables...")
    try:
        df_par = load_file(parallel_path)
        print(f" -> Loaded {len(df_par)} records from Parallel Run.")
    except Exception as e:
        print(f"Error loading Parallel Run file: {e}")
        return

    print("Cleaning APL Policy Numbers...")
    df_apl['Cleaned Policy Number'] = df_apl[APLColumns.POLICY_NUMBER].str.replace("'", "", regex=False)
    # Remove OVR prefix or suffix (case insensitive)
    df_apl['Cleaned Policy Number'] = df_apl['Cleaned Policy Number'].str.replace(r'(?i)^OVR[-_ ]?|[-_ ]?OVR$', '', regex=True)

    print("Building Parallel Run lookup pool for 1-to-1 matching...")
    par_pool = {}
    for idx, row in df_par.iterrows():
        pol = str(row[ParallelColumns.POLICY_NUMBER])
        if pol not in par_pool:
            par_pool[pol] = []
        par_pool[pol].append({'row': row, 'consumed': False, 'idx': idx})
        
    parallel_agent_col = ParallelColumns.AGENT_NAME + '_parallel' if ParallelColumns.AGENT_NAME in df_apl.columns else ParallelColumns.AGENT_NAME
    
    results = []
    passed_results = []
    stats = {
        'total_processed': len(df_apl),
        'exact_matches': 0,
        'partial_matches': 0,
        'missing_in_parallel': 0,
        'agent_name_mismatch': 0,
        'amount_mismatch': 0,
        'extra_in_parallel': 0
    }
    
    print("Processing matching and comparisons optimally...")
    
    # Prepare APL items and group them by target parallel policy number
    apl_items = []
    for idx, row in df_apl.iterrows():
        apl_cleaned = str(row['Cleaned Policy Number'])
        target_pol = None
        is_partial = False
        
        if apl_cleaned in par_pool:
            target_pol = apl_cleaned
        elif apl_cleaned and apl_cleaned[:-1] in par_pool:
            target_pol = apl_cleaned[:-1]
            is_partial = True
            
        apl_items.append({
            'row': row,
            'idx': idx,
            'apl_cleaned': apl_cleaned,
            'target_pol': target_pol,
            'is_partial': is_partial,
            'matched_par': None,
            'match_type': "None"
        })

    # Perform multi-pass matching for each policy group
    for target_pol, par_list in par_pool.items():
        target_apl_items = [item for item in apl_items if item['target_pol'] == target_pol]
        if not target_apl_items:
            continue
            
        # Pass 1: Exact matches (Agent + Amount)
        for apl_item in target_apl_items:
            if apl_item['matched_par']: continue
            apl_row = apl_item['row']
            apl_agent = str(apl_row.get(APLColumns.AGENT_NAME, "")).strip()
            if apl_agent.lower() == 'nan': apl_agent = ""
            
            for par_item in par_list:
                if par_item['consumed']: continue
                icm_ag = str(par_item['row'].get(ParallelColumns.AGENT_NAME, "")).strip()
                if icm_ag.lower() == 'nan': icm_ag = ""
                icm_amt_str = str(par_item['row'].get(ParallelColumns.AMOUNT, "")).strip()
                if icm_amt_str.lower() == 'nan': icm_amt_str = ""
                icm_amt_val = clean_amount(icm_amt_str)
                
                ag_match = normalize_name(apl_agent) == normalize_name(icm_ag)
                matched_col, _ = check_amount_match(apl_row, icm_amt_val, amount_tolerance)
                amt_match = matched_col != "None"
                
                if ag_match and amt_match:
                    apl_item['matched_par'] = par_item
                    apl_item['match_type'] = "Partial" if apl_item['is_partial'] else "Exact"
                    par_item['consumed'] = True
                    break
                    
        # Pass 2: Agent matches
        for apl_item in target_apl_items:
            if apl_item['matched_par']: continue
            apl_row = apl_item['row']
            apl_agent = str(apl_row.get(APLColumns.AGENT_NAME, "")).strip()
            if apl_agent.lower() == 'nan': apl_agent = ""
            
            for par_item in par_list:
                if par_item['consumed']: continue
                icm_ag = str(par_item['row'].get(ParallelColumns.AGENT_NAME, "")).strip()
                if icm_ag.lower() == 'nan': icm_ag = ""
                
                if normalize_name(apl_agent) == normalize_name(icm_ag):
                    apl_item['matched_par'] = par_item
                    apl_item['match_type'] = "Partial" if apl_item['is_partial'] else "Exact"
                    par_item['consumed'] = True
                    break
                    
        # Pass 3: Amount matches
        for apl_item in target_apl_items:
            if apl_item['matched_par']: continue
            apl_row = apl_item['row']
            
            for par_item in par_list:
                if par_item['consumed']: continue
                icm_amt_str = str(par_item['row'].get(ParallelColumns.AMOUNT, "")).strip()
                if icm_amt_str.lower() == 'nan': icm_amt_str = ""
                icm_amt_val = clean_amount(icm_amt_str)
                
                matched_col, _ = check_amount_match(apl_row, icm_amt_val, amount_tolerance)
                if matched_col != "None":
                    apl_item['matched_par'] = par_item
                    apl_item['match_type'] = "Partial" if apl_item['is_partial'] else "Exact"
                    par_item['consumed'] = True
                    break
                    
        # Pass 4: Arbitrary matches for remaining
        for apl_item in target_apl_items:
            if apl_item['matched_par']: continue
            
            for par_item in par_list:
                if par_item['consumed']: continue
                apl_item['matched_par'] = par_item
                apl_item['match_type'] = "Partial" if apl_item['is_partial'] else "Exact"
                par_item['consumed'] = True
                break

    # Build results
    for apl_item in apl_items:
        apl_row = apl_item['row']
        apl_pol = str(apl_row[APLColumns.POLICY_NUMBER]) if pd.notnull(apl_row[APLColumns.POLICY_NUMBER]) else ""
        apl_agent = str(apl_row[APLColumns.AGENT_NAME]).strip() if pd.notnull(apl_row[APLColumns.AGENT_NAME]) else ""
        if apl_agent.lower() == 'nan': apl_agent = ""
        apl_product = str(apl_row.get(APLColumns.PRODUCT, "")).strip()
        if apl_product.lower() == 'nan': apl_product = ""
        
        best_match = apl_item['matched_par']
        match_type = apl_item['match_type']
        
        if best_match:
            icm_row = best_match['row']
            icm_pol = str(icm_row[ParallelColumns.POLICY_NUMBER]) if pd.notnull(icm_row[ParallelColumns.POLICY_NUMBER]) else ""
            if icm_pol.lower() == 'nan': icm_pol = ""
            
            icm_agent = str(icm_row.get(ParallelColumns.AGENT_NAME, "")).strip()
            if icm_agent.lower() == 'nan': icm_agent = ""
            
            icm_amount_str = str(icm_row.get(ParallelColumns.AMOUNT, "")).strip()
            if icm_amount_str.lower() == 'nan': icm_amount_str = ""
            
            icm_product = str(icm_row.get(ParallelColumns.PRODUCT, "")).strip()
            if icm_product.lower() == 'nan': icm_product = ""
        else:
            icm_pol = ""
            icm_agent = ""
            icm_amount_str = ""
            icm_product = ""
            
        errors = []
        if match_type == "Partial":
            errors.append("Partial Match (APL has extra last character)")
            stats['partial_matches'] += 1

        if not icm_pol:
            errors.append("Policy not found in Parallel Run")
            stats['missing_in_parallel'] += 1
            matched_col = "None"
            apl_amount_str = str(apl_row.get("Earned Commission", "")).strip()
            if apl_amount_str.lower() == 'nan': apl_amount_str = ""
        else:
            if normalize_name(apl_agent) != normalize_name(icm_agent):
                errors.append("Agent name does not match")
                stats['agent_name_mismatch'] += 1
            
            icm_amt_val = clean_amount(icm_amount_str)
            matched_col, apl_amount_str = check_amount_match(apl_row, icm_amt_val, amount_tolerance)
            if matched_col == "None":
                apl_amt_val = clean_amount(apl_amount_str)
                errors.append(f"Amount mismatch (APL: {apl_amt_val}, ICM: {icm_amt_val})")
                stats['amount_mismatch'] += 1
                
            if not errors:
                stats['exact_matches'] += 1
                
        row_dict = {
            'APL Policy Number': apl_pol,
            'APL Agent Name': apl_agent,
            'APL Product': apl_product,
            'APL Amount': apl_amount_str,
            'ICM Policy Number': icm_pol,
            'ICM Agent Name': icm_agent,
            'ICM Product': icm_product,
            'ICM Amount': icm_amount_str,
            'Matched APL Amount Column': matched_col,
            'Error': " | ".join(errors)
        }
        
        if not errors or (len(errors) == 1 and errors[0].startswith("Partial Match")):
            passed_results.append(row_dict)
        else:
            results.append(row_dict)

    print("Sweeping unconsumed Parallel Run rows...")
    for pol, pool_list in par_pool.items():
        for item in pool_list:
            if not item['consumed']:
                icm_row = item['row']
                icm_pol = str(icm_row[ParallelColumns.POLICY_NUMBER]) if pd.notnull(icm_row[ParallelColumns.POLICY_NUMBER]) else ""
                icm_agent = str(icm_row.get(ParallelColumns.AGENT_NAME, "")).strip()
                icm_amount_str = str(icm_row.get(ParallelColumns.AMOUNT, "")).strip()
                icm_product = str(icm_row.get(ParallelColumns.PRODUCT, "")).strip()
                
                results.append({
                    'APL Policy Number': "",
                    'APL Agent Name': "",
                    'APL Product': "",
                    'APL Amount': "",
                    'ICM Policy Number': icm_pol,
                    'ICM Agent Name': icm_agent,
                    'ICM Product': icm_product,
                    'ICM Amount': icm_amount_str,
                    'Matched APL Amount Column': "",
                    'Error': "Found in Parallel Run but missing from APL Statement"
                })
                stats['extra_in_parallel'] += 1
        
    out_df = pd.DataFrame(results)
    
    total = stats['total_processed']
    def pct(val):
        return f"({(val/total)*100:.1f}%)" if total > 0 else "(0.0%)"
        
    print("\n--- Validation Summary ---")
    print(f"Total Records Processed: {total}")
    print(f"Exact Matches:           {stats['exact_matches']:<6} {pct(stats['exact_matches'])}")
    print(f"Partial Matches (1 char):{stats['partial_matches']:<6} {pct(stats['partial_matches'])}")
    print(f"Missing in Parallel:     {stats['missing_in_parallel']:<6} {pct(stats['missing_in_parallel'])}")
    print(f"Extra in Parallel:       {stats['extra_in_parallel']:<6} (Orphaned ICM rows)")
    print(f"Agent Name Mismatches:   {stats['agent_name_mismatch']:<6} {pct(stats['agent_name_mismatch'])}")
    print(f"Amount Mismatches:       {stats['amount_mismatch']:<6} {pct(stats['amount_mismatch'])}")
    print("--------------------------\n")
    
    print(f"Saving mismatches to {out_csv_path} and {out_xlsx_path}...")
    out_df.to_csv(out_csv_path, index=False)
    out_df.to_excel(out_xlsx_path, index=False)
    
    # Segregate errors into separate files
    errors_dir = os.path.join(os.path.dirname(out_csv_path), 'Segregated_Errors')
    os.makedirs(errors_dir, exist_ok=True)
    
    def categorize_error(err_str):
        if not isinstance(err_str, str): return "Unknown"
        if "Found in Parallel Run but missing" in err_str: return "Extra_in_Parallel"
        if "Policy not found in Parallel Run" in err_str: return "Missing_in_Parallel"
        
        has_agent = "Agent name does not match" in err_str
        has_amt = "Amount mismatch" in err_str
        
        if has_agent and has_amt: return "Agent_and_Amount_Mismatch"
        if has_agent: return "Agent_Name_Mismatch"
        if has_amt: return "Amount_Mismatch"
        return "Other_Errors"
        
    print(f"Segregating errors into folder: {errors_dir}...")
    if not out_df.empty:
        out_df['Error Category'] = out_df['Error'].apply(categorize_error)
        for cat, group in out_df.groupby('Error Category'):
            group.drop(columns=['Error Category']).to_csv(os.path.join(errors_dir, f"{cat}.csv"), index=False)
    
    out_passed_df = pd.DataFrame(passed_results)
    if not out_passed_df.empty:
        passed_csv_path = out_csv_path.replace('Validation_Output', 'Passed_Output')
        passed_xlsx_path = out_xlsx_path.replace('Validation_Output', 'Passed_Output')
        print(f"Saving passed matches to {passed_csv_path}...")
        out_passed_df.to_csv(passed_csv_path, index=False)
        out_passed_df.to_excel(passed_xlsx_path, index=False)
        
    print("Done!")

if __name__ == "__main__":
    print("Welcome to the Policy Commission Validation Script.")
    apl_file = input("Enter path to APL Statement CSV: ").strip().strip('\'"')
    parallel_file = input("Enter path to Parallel Run Payables CSV: ").strip().strip('\'"')
    
    if not (os.path.exists(apl_file) and os.path.exists(parallel_file)):
        print("Error: One or more input files do not exist. Please check the paths and try again.")
        exit(1)
        
    tolerance = 0.01
    use_tol = input("Do you want to set a custom amount tolerance? (y/n) [default: n]: ").strip().lower()
    if use_tol in ['y', 'yes']:
        tol_input = input("Enter amount tolerance (e.g. 0.05, 1.5): ").strip()
        try:
            tolerance = float(tol_input)
        except ValueError:
            print("Invalid input, defaulting to 0.01")
            tolerance = 0.01
        
    # Determine next version folder
    base_dir = os.path.dirname(os.path.abspath(__file__))
    versions_dir = os.path.join(base_dir, "versions")
    os.makedirs(versions_dir, exist_ok=True)
    
    existing_versions = [d for d in os.listdir(versions_dir) if d.startswith('v') and os.path.isdir(os.path.join(versions_dir, d))]
    next_v = 1
    if existing_versions:
        v_nums = [int(v[1:]) for v in existing_versions if v[1:].isdigit()]
        if v_nums:
            next_v = max(v_nums) + 1
            
    version_folder = os.path.join(versions_dir, f"v{next_v}")
    input_folder = os.path.join(version_folder, "input")
    output_folder = os.path.join(version_folder, "output")
    
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    
    print(f"\nCreating version folder: {version_folder}")
    
    # Copy files
    print("Copying input files to version input folder...")
    shutil.copy2(apl_file, os.path.join(input_folder, os.path.basename(apl_file)))
    shutil.copy2(parallel_file, os.path.join(input_folder, os.path.basename(parallel_file)))
    
    out_csv = os.path.join(output_folder, 'Validation_Output.csv')
    out_xlsx = os.path.join(output_folder, 'Validation_Output.xlsx')
    
    print("Starting validation...\n")
    validate_data(apl_file, parallel_file, out_csv, out_xlsx, tolerance)
    print(f"\nValidation complete. Outputs saved in {output_folder}")
