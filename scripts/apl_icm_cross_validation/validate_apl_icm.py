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
import argparse

VALID_MODES = ("apl", "icm")


def normalize_mode(mode):
    if mode is None:
        return None
    value = str(mode).strip().lower()
    aliases = {
        "1": "apl",
        "apl": "apl",
        "apl_to_icm": "apl",
        "apl->icm": "apl",
        "2": "icm",
        "icm": "icm",
        "icm_to_apl": "icm",
        "icm->apl": "icm",
        "parallel": "icm",
    }
    return aliases.get(value, value)


def prompt_for_mode():
    print("Select validation mode:")
    print("  1) APL -> ICM  (drive from APL statement) [default]")
    print("  2) ICM -> APL  (drive from forecast/ICM payables)")
    choice = input("Enter mode [1/2 or apl/icm] (default: 1): ").strip() or "1"
    mode = normalize_mode(choice)
    if mode not in VALID_MODES:
        print(f"Invalid mode '{choice}', defaulting to apl")
        return "apl"
    return mode


def normalize_name(name):
    if pd.isna(name): return set()
    # Normalize punctuation (convert hyphens to space, remove commas/dots) and collapse multiple spaces
    name_str = re.sub(r'\s+', ' ', re.sub(r'[,.]', '', str(name).replace('-', ' ')).lower()).strip()
    
    # Alias Main Line Benefits / MLB / MLB Parallel Run
    if 'main line' in name_str or 'mlb' in name_str:
        return {'mlb'}
        
    words = name_str.split()
    if not words: return set()

    # Filter out common name suffixes (IV, III, II, JR, SR, etc.)
    suffixes = {'iv', 'iii', 'ii', 'v', 'vi', 'jr', 'sr', 'esq', '1st', '2nd', '3rd', '4th'}
    words = [w for w in words if w not in suffixes]
    if not words: return set()

    nickname_map = {
        'greg': 'gregory',
    }
    words = [nickname_map.get(w, w) for w in words]

    if len(words) == 1: return {words[0]}
    return {words[0], words[-1]}

def clean_policy_number(val):
    if val is None or pd.isna(val):
        return ""
    v_str = str(val).strip()
    # Drop quote marks anywhere (APL values often look like '00000825606A8).
    v_str = v_str.replace("'", "").replace('"', "").strip()
    # Override trailer: OVR, -OVR, _OVR, KOVR, K_BOVR, and the same shapes with a leading letter.
    v_str = re.sub(r'(?i)[A-Z]?(?:[_ -][A-Z]*)?OVR$', '', v_str).strip()
    v_str = re.sub(r'(?i)^OVR[_ -]?', '', v_str).strip()
    if v_str.endswith(".0"):
        v_str = v_str[:-2].strip()
    # ICM often pads with a leading zero (055733973) while APL stores 55733973.
    stripped = v_str.lstrip("0")
    return stripped or ("0" if v_str else "")

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

def _amount_cell(row, col):
    if col not in row:
        return "", 0.0
    val_str = str(row[col]).strip() if pd.notnull(row[col]) else ""
    if val_str.lower() == 'nan':
        val_str = ""
    return val_str, clean_amount(val_str)


def apl_line_amount(row):
    """Line amount used for pairing: first non-zero money column, otherwise zero."""
    zero_hit = None
    for col in AMOUNT_COLUMNS:
        val_str, val_num = _amount_cell(row, col)
        if not val_str and val_num == 0.0:
            continue
        if abs(val_num) > 0:
            return col, val_str, val_num
        if zero_hit is None:
            zero_hit = (col, val_str, 0.0)
    if zero_hit is not None:
        return zero_hit
    return "Earned Commission", "", 0.0


def check_amount_match(row, icm_amt_val, tolerance=0.01):
    # A $0 Earned Credit on a row that already has Earned Commission must not
    # consume an ICM $0 split. Only a row whose own line amount is ~0 can match $0.
    col, val_str, val_num = apl_line_amount(row)
    if abs(val_num - icm_amt_val) <= tolerance:
        return col, val_str if val_str else f"{val_num}"
    return "None", val_str


def is_policy_suffix(suffix):
    """APL trailers seen on statements: one extra character, or a short digit run (2, 26)."""
    if not suffix or len(suffix) > 4:
        return False
    if len(suffix) == 1:
        return True
    return suffix.isdigit()


def match_policy_base(apl_cleaned, pool):
    """Return (icm policy, is_partial). Partial means the APL number has a trailer on a known policy."""
    if not apl_cleaned:
        return None, False
    if apl_cleaned in pool:
        return apl_cleaned, False
    best = None
    for key in pool:
        if not key or len(key) >= len(apl_cleaned) or not apl_cleaned.startswith(key):
            continue
        if not is_policy_suffix(apl_cleaned[len(key):]):
            continue
        if best is None or len(key) > len(best):
            best = key
    if best is not None:
        return best, True
    return None, False

def align_apl_columns(df):
    """Map carrier statement headers onto the names this checker expects."""
    normalized = {str(column).strip().lower(): column for column in df.columns}
    aliases = {
        "Policy Number": ["policy number", "policy no", "policy no.", "policynumber", "policy_number"],
        "Agent Name": ["agent name", "writing agent name", "name"],
        "Product": ["product", "product name"],
    }
    rename = {}
    for expected, candidates in aliases.items():
        if expected in df.columns:
            continue
        for candidate in candidates:
            source = normalized.get(candidate)
            if source is not None and source not in rename:
                rename[source] = expected
                break
    if rename:
        df = df.rename(columns=rename)
    if "Product" not in df.columns:
        df["Product"] = ""
    missing = [name for name in ("Policy Number", "Agent Name") if name not in df.columns]
    if missing:
        raise KeyError(
            "APL statement is missing required column(s): "
            + ", ".join(missing)
            + ". Found columns: "
            + ", ".join(str(column) for column in df.columns)
        )
    return df


def load_file(file_path):
    if file_path.lower().endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path, dtype=str)
    else:
        try:
            return pd.read_csv(file_path, dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            return pd.read_csv(file_path, dtype=str, encoding='latin-1')

def extract_meta_val(row, default_filename, *keys):
    if row is None:
        return default_filename if ("file" in keys[0].lower() and default_filename) else ""
    if hasattr(row, 'index'):
        col_map = {str(c).strip().lower().replace("_", "").replace(" ", ""): c for c in row.index}
        for k in keys:
            norm_k = k.lower().replace("_", "").replace(" ", "")
            if norm_k in col_map:
                val = str(row[col_map[norm_k]]).strip()
                if val and val.lower() != 'nan':
                    return val
    return default_filename if ("file" in keys[0].lower() and default_filename) else ""


def _text(row, column):
    if row is None or column not in getattr(row, "index", []):
        return ""
    if pd.isna(row[column]):
        return ""
    text = str(row[column]).strip()
    return "" if text.lower() == "nan" else text


def _policy_text(row, column):
    return _text(row, column).replace("'", "").replace('"', "").strip()


def pair_apl_and_icm(df_apl, df_par, amount_tolerance):
    """Pair APL lines to ICM lines with the same four passes validate_data uses."""
    print("Cleaning APL Policy Numbers...")
    df_apl['Cleaned Policy Number'] = df_apl[APLColumns.POLICY_NUMBER].apply(clean_policy_number)

    print("Building Parallel Run lookup pool for 1-to-1 matching...")
    par_pool = {}
    for idx, row in df_par.iterrows():
        pol = clean_policy_number(row[ParallelColumns.POLICY_NUMBER])
        if pol not in par_pool:
            par_pool[pol] = []
        par_pool[pol].append({'row': row, 'consumed': False, 'idx': idx})

    print("Processing matching and comparisons optimally...")

    apl_items = []
    for idx, row in df_apl.iterrows():
        apl_cleaned = str(row['Cleaned Policy Number'])
        target_pol, is_partial = match_policy_base(apl_cleaned, par_pool)
        apl_items.append({
            'row': row,
            'idx': idx,
            'apl_cleaned': apl_cleaned,
            'target_pol': target_pol,
            'is_partial': is_partial,
            'matched_par': None,
            'match_type': "None"
        })

    def bind_match(apl_item, par_item):
        apl_item['matched_par'] = par_item
        apl_item['match_type'] = "Partial" if apl_item['is_partial'] else "Exact"
        par_item['consumed'] = True
        par_item['matched_apl'] = apl_item

    for target_pol, par_list in par_pool.items():
        target_apl_items = [item for item in apl_items if item['target_pol'] == target_pol]
        if not target_apl_items:
            continue

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
                    bind_match(apl_item, par_item)
                    break

        for apl_item in target_apl_items:
            if apl_item['matched_par']: continue
            apl_row = apl_item['row']
            apl_agent = str(apl_row.get(APLColumns.AGENT_NAME, "")).strip()
            if apl_agent.lower() == 'nan': apl_agent = ""
            _, _, apl_amt = apl_line_amount(apl_row)

            best_par = None
            best_diff = None
            for par_item in par_list:
                if par_item['consumed']: continue
                icm_ag = str(par_item['row'].get(ParallelColumns.AGENT_NAME, "")).strip()
                if icm_ag.lower() == 'nan': icm_ag = ""
                if normalize_name(apl_agent) != normalize_name(icm_ag):
                    continue
                icm_amt_str = str(par_item['row'].get(ParallelColumns.AMOUNT, "")).strip()
                if icm_amt_str.lower() == 'nan': icm_amt_str = ""
                diff = abs(apl_amt - clean_amount(icm_amt_str))
                if best_par is None or diff < best_diff:
                    best_par = par_item
                    best_diff = diff
            if best_par is not None:
                bind_match(apl_item, best_par)

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
                    bind_match(apl_item, par_item)
                    break

        for apl_item in target_apl_items:
            if apl_item['matched_par']: continue

            for par_item in par_list:
                if par_item['consumed']: continue
                bind_match(apl_item, par_item)
                break

    return apl_items, par_pool


def compare_icm_to_apl(apl_path, icm_path, amount_tolerance=0.01):
    """Return one record per ICM line, plus APL lines that never paired."""
    print("Loading APL Statement...")
    df_apl = align_apl_columns(load_file(apl_path))
    print(f" -> Loaded {len(df_apl)} records from APL.")

    print("Loading ICM payables...")
    df_par = load_file(icm_path)
    print(f" -> Loaded {len(df_par)} records from ICM.")
    missing = [
        name for name in (ParallelColumns.POLICY_NUMBER, ParallelColumns.AGENT_NAME, ParallelColumns.AMOUNT)
        if name not in df_par.columns
    ]
    if missing:
        raise KeyError(
            "ICM payables are missing required column(s): "
            + ", ".join(missing)
            + ". Found columns: "
            + ", ".join(str(column) for column in df_par.columns)
        )

    apl_items, par_pool = pair_apl_and_icm(df_apl, df_par, amount_tolerance)

    icm_lines = []
    for pool_list in par_pool.values():
        for item in pool_list:
            icm_row = item['row']
            matched = item.get('matched_apl')
            apl_row = matched['row'] if matched else None
            apl_amount = apl_line_amount(apl_row)[2] if apl_row is not None else 0.0
            icm_lines.append({
                'icm_policy': _policy_text(icm_row, ParallelColumns.POLICY_NUMBER),
                'icm_agent': _text(icm_row, ParallelColumns.AGENT_NAME),
                'icm_agent_id': extract_meta_val(
                    icm_row, "", "agent_id", "agentid", "agent id",
                    "writing_agent_id", "writing agent id", "agent_code", "agent code",
                ),
                'icm_product': _text(icm_row, ParallelColumns.PRODUCT),
                'icm_commission_type': extract_meta_val(
                    icm_row, "", "commission_type", "commissiontype", "commission type",
                    "commission_type_text", "commission type text",
                ),
                'icm_amount': round(clean_amount(_text(icm_row, ParallelColumns.AMOUNT)), 2),
                'icm_statement_file': extract_meta_val(
                    icm_row, "",
                    "source_file_name", "sourcefilename", "file_name", "filename", "source file name",
                ),
                'paired': apl_row is not None,
                'apl_policy': _policy_text(apl_row, APLColumns.POLICY_NUMBER) if apl_row is not None else "",
                'apl_agent': _text(apl_row, APLColumns.AGENT_NAME) if apl_row is not None else "",
                'apl_product': _text(apl_row, APLColumns.PRODUCT) if apl_row is not None else "",
                'apl_amount': round(apl_amount, 2) if apl_row is not None else 0.0,
            })

    apl_only = []
    for apl_item in apl_items:
        if apl_item['matched_par']:
            continue
        apl_row = apl_item['row']
        apl_only.append({
            'apl_policy': _policy_text(apl_row, APLColumns.POLICY_NUMBER),
            'apl_agent': _text(apl_row, APLColumns.AGENT_NAME),
            'apl_product': _text(apl_row, APLColumns.PRODUCT),
            'apl_amount': round(apl_line_amount(apl_row)[2], 2),
        })
    return icm_lines, apl_only


def validate_data(apl_path, parallel_path, out_csv_path, out_xlsx_path, amount_tolerance=0.01, mode="apl", single_grouped_file=False, apl_df=None):
    mode = normalize_mode(mode) or "apl"
    if mode not in VALID_MODES:
        raise ValueError(f"Unsupported validation mode '{mode}'. Expected one of: {', '.join(VALID_MODES)}")
    print(f"Validation mode: {mode} ({'APL drives' if mode == 'apl' else 'ICM drives'})")

    apl_filename = os.path.basename(apl_path)
    parallel_filename = os.path.basename(parallel_path)

    print("Loading APL Statement...")
    try:
        if apl_df is not None:
            df_apl = apl_df.copy()
            print(f" -> Reused {len(df_apl)} cached APL records.")
        else:
            df_apl = align_apl_columns(load_file(apl_path))
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

    results = []
    passed_results = []
    stats = {
        'total_processed': len(df_par) if mode == "icm" else len(df_apl),
        'exact_matches': 0,
        'partial_matches': 0,
        'missing_in_parallel': 0,
        'agent_name_mismatch': 0,
        'amount_mismatch': 0,
        'extra_in_parallel': 0,
        'zero_splits_not_on_apl': 0,
    }

    apl_items, par_pool = pair_apl_and_icm(df_apl, df_par, amount_tolerance)

    def append_comparison_row(apl_row, icm_row, match_type, missing_message):
        apl_pol = ""
        apl_agent = ""
        apl_product = ""
        if apl_row is not None:
            apl_pol = str(apl_row[APLColumns.POLICY_NUMBER]) if pd.notnull(apl_row[APLColumns.POLICY_NUMBER]) else ""
            apl_agent = str(apl_row[APLColumns.AGENT_NAME]).strip() if pd.notnull(apl_row[APLColumns.AGENT_NAME]) else ""
            if apl_agent.lower() == 'nan':
                apl_agent = ""
            apl_product = str(apl_row.get(APLColumns.PRODUCT, "")).strip()
            if apl_product.lower() == 'nan':
                apl_product = ""

        if icm_row is not None:
            icm_pol = str(icm_row[ParallelColumns.POLICY_NUMBER]) if pd.notnull(icm_row[ParallelColumns.POLICY_NUMBER]) else ""
            if icm_pol.lower() == 'nan':
                icm_pol = ""
            icm_agent = str(icm_row.get(ParallelColumns.AGENT_NAME, "")).strip()
            if icm_agent.lower() == 'nan':
                icm_agent = ""
            icm_amount_str = str(icm_row.get(ParallelColumns.AMOUNT, "")).strip()
            if icm_amount_str.lower() == 'nan':
                icm_amount_str = ""
            icm_product = str(icm_row.get(ParallelColumns.PRODUCT, "")).strip()
            if icm_product.lower() == 'nan':
                icm_product = ""
        else:
            icm_pol = ""
            icm_agent = ""
            icm_amount_str = ""
            icm_product = ""

        errors = []
        if match_type == "Partial":
            # A trailer such as 2 or 26 on the APL policy is still the same policy.
            stats['partial_matches'] += 1

        if missing_message:
            errors.append(missing_message)
            if "Parallel Run" in missing_message and "not found" in missing_message:
                stats['missing_in_parallel'] += 1
            elif "APL Statement" in missing_message and "not found" in missing_message:
                stats['extra_in_parallel'] += 1
            matched_col = "None"
            apl_amount_str = ""
            if apl_row is not None:
                apl_amount_str = str(apl_row.get("Earned Commission", "")).strip()
                if apl_amount_str.lower() == 'nan':
                    apl_amount_str = ""
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

            if not errors and match_type != "Partial":
                stats['exact_matches'] += 1

        tx_id = extract_meta_val(icm_row, "", "transaction_id", "transactionid", "transaction id") or extract_meta_val(apl_row, "", "transaction_id", "transactionid", "transaction id")
        batch_id = extract_meta_val(icm_row, "", "source_batch_id", "sourcebatchid", "batch_id", "batchid", "source batch id") or extract_meta_val(apl_row, "", "source_batch_id", "sourcebatchid", "batch_id", "batchid", "source batch id")
        file_name = extract_meta_val(icm_row, parallel_filename, "source_file_name", "sourcefilename", "file_name", "filename", "source file name") or extract_meta_val(apl_row, apl_filename, "source_file_name", "sourcefilename", "file_name", "filename", "source file name")

        row_dict = {
            'transaction_id': tx_id,
            'source_batch_id': batch_id,
            'source_file_name': file_name,
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

    if mode == "apl":
        # APL drives: emit one row per APL record.
        for apl_item in apl_items:
            best_match = apl_item['matched_par']
            append_comparison_row(
                apl_item['row'],
                best_match['row'] if best_match else None,
                apl_item['match_type'],
                None if best_match else "Policy not found in Parallel Run",
            )
    elif mode == "icm":
        # ICM drives: emit one row per ICM record.
        print("Building ICM-driven comparison rows...")
        for _, pool_list in par_pool.items():
            for item in pool_list:
                matched_apl = item.get('matched_apl')
                if not matched_apl:
                    icm_amt_str = str(item['row'].get(ParallelColumns.AMOUNT, "")).strip()
                    if icm_amt_str.lower() == 'nan':
                        icm_amt_str = ""
                    # APL omits agents who were paid $0. That is not a payable gap.
                    if abs(clean_amount(icm_amt_str)) <= amount_tolerance:
                        stats['zero_splits_not_on_apl'] += 1
                        continue
                append_comparison_row(
                    matched_apl['row'] if matched_apl else None,
                    item['row'],
                    matched_apl['match_type'] if matched_apl else "None",
                    None if matched_apl else "Policy not found in APL Statement",
                )
    else:
        # Full bidirectional mode: APL drives + sweep unconsumed ICM rows + sweep unconsumed APL rows
        for apl_item in apl_items:
            best_match = apl_item['matched_par']
            append_comparison_row(
                apl_item['row'],
                best_match['row'] if best_match else None,
                apl_item['match_type'],
                None if best_match else "Policy not found in Parallel Run",
            )

        print("Sweeping unconsumed Parallel Run rows...")
        for _, pool_list in par_pool.items():
            for item in pool_list:
                if item['consumed']:
                    continue
                icm_row = item['row']
                icm_pol = str(icm_row[ParallelColumns.POLICY_NUMBER]) if pd.notnull(icm_row[ParallelColumns.POLICY_NUMBER]) else ""
                icm_agent = str(icm_row.get(ParallelColumns.AGENT_NAME, "")).strip()
                icm_amount_str = str(icm_row.get(ParallelColumns.AMOUNT, "")).strip()
                icm_product = str(icm_row.get(ParallelColumns.PRODUCT, "")).strip()

                tx_id = extract_meta_val(icm_row, "", "transaction_id", "transactionid", "transaction id")
                batch_id = extract_meta_val(icm_row, "", "source_batch_id", "sourcebatchid", "batch_id", "batchid", "source batch id")
                file_name = extract_meta_val(icm_row, parallel_filename, "source_file_name", "sourcefilename", "file_name", "filename", "source file name")

                results.append({
                    'transaction_id': tx_id,
                    'source_batch_id': batch_id,
                    'source_file_name': file_name,
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

        print("Sweeping unconsumed APL Statement rows...")
        for apl_item in apl_items:
            if apl_item['matched_par']:
                continue
            apl_row = apl_item['row']
            apl_pol = str(apl_row[APLColumns.POLICY_NUMBER]) if pd.notnull(apl_row[APLColumns.POLICY_NUMBER]) else ""
            apl_agent = str(apl_row[APLColumns.AGENT_NAME]).strip() if pd.notnull(apl_row[APLColumns.AGENT_NAME]) else ""
            apl_product = str(apl_row.get(APLColumns.PRODUCT, "")).strip()
            apl_amount_str = str(apl_row.get("Earned Commission", "")).strip()
            if apl_amount_str.lower() == 'nan':
                apl_amount_str = ""

            tx_id = extract_meta_val(apl_row, "", "transaction_id", "transactionid", "transaction id")
            batch_id = extract_meta_val(apl_row, "", "source_batch_id", "sourcebatchid", "batch_id", "batchid", "source batch id")
            file_name = extract_meta_val(apl_row, apl_filename, "source_file_name", "sourcefilename", "file_name", "filename", "source file name")

            results.append({
                'transaction_id': tx_id,
                'source_batch_id': batch_id,
                'source_file_name': file_name,
                'APL Policy Number': apl_pol,
                'APL Agent Name': apl_agent,
                'APL Product': apl_product,
                'APL Amount': apl_amount_str,
                'ICM Policy Number': "",
                'ICM Agent Name': "",
                'ICM Product': "",
                'ICM Amount': "",
                'Matched APL Amount Column': "",
                'Error': "Found in APL Statement but missing from Parallel Run"
            })
            stats['missing_in_parallel'] += 1
        
    out_df = pd.DataFrame(results)
    
    total = stats['total_processed']
    def pct(val):
        return f"({(val/total)*100:.1f}%)" if total > 0 else "(0.0%)"
        
    print("\n--- Validation Summary ---")
    print(f"Total Records Processed: {total}")
    print(f"Exact Matches:           {stats['exact_matches']:<6} {pct(stats['exact_matches'])}")
    print(f"Policy suffix matches:   {stats['partial_matches']:<6} {pct(stats['partial_matches'])}")
    print(f"Missing in Parallel:     {stats['missing_in_parallel']:<6} {pct(stats['missing_in_parallel'])}")
    print(f"Extra in Parallel:       {stats['extra_in_parallel']:<6} (Orphaned ICM rows)")
    print(f"Zero splits not on APL:  {stats['zero_splits_not_on_apl']:<6} {pct(stats['zero_splits_not_on_apl'])}")
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
        if "Policy not found in APL Statement" in err_str: return "Extra_in_Parallel"
        if "Policy not found in Parallel Run" in err_str: return "Missing_in_Parallel"
        if "Found in APL Statement but missing from Parallel Run" in err_str: return "Missing_in_Parallel"
        
        has_agent = "Agent name does not match" in err_str
        has_amt = "Amount mismatch" in err_str
        
        if has_agent and has_amt: return "Agent_and_Amount_Mismatch"
        if has_agent: return "Agent_Name_Mismatch"
        if has_amt: return "Amount_Mismatch"
        return "Other_Errors"
        
    print(f"Segregating errors into folder: {errors_dir}...")
    for stale_name in os.listdir(errors_dir):
        stale_path = os.path.join(errors_dir, stale_name)
        if os.path.isfile(stale_path):
            os.remove(stale_path)
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

    # ---------------------------------------------------------------------------
    # ADDITIONAL GROUPED POLICY FILES
    # Grouped by (source_batch_id, source_file_name, base_policy_number).
    # If ANY row in a policy group has an error/failure, the ENTIRE policy group moves to Grouped_Policy_Failed.
    # Only policy groups where 100% of rows passed move to Grouped_Policy_Passed.
    # ---------------------------------------------------------------------------
    output_dir = os.path.dirname(out_csv_path)

    all_processed_rows = passed_results + results
    if all_processed_rows:
        def clean_pol_key(val):
            return clean_policy_number(val)

        def get_group_key(r):
            batch_id = str(r.get('source_batch_id') or '').strip()
            file_name = str(r.get('source_file_name') or '').strip()
            icm_pol = clean_pol_key(r.get('ICM Policy Number'))
            apl_pol = clean_pol_key(r.get('APL Policy Number'))
            pol_no = icm_pol or apl_pol
            return (batch_id, file_name, pol_no)

        # Identify policy groups that contain AT LEAST ONE row with errors / failures
        failed_group_keys = set()
        for r in all_processed_rows:
            group_key = get_group_key(r)
            err_str = str(r.get('Error') or '').strip()
            is_error = bool(err_str and not err_str.startswith("Partial Match"))
            if is_error or r in results:
                failed_group_keys.add(group_key)

        grouped_passed_rows = []
        grouped_failed_rows = []

        for r in all_processed_rows:
            group_key = get_group_key(r)
            if group_key in failed_group_keys:
                grouped_failed_rows.append(r)
            else:
                grouped_passed_rows.append(r)

        grp_passed_csv = os.path.join(output_dir, 'Grouped_Policy_Passed.csv')
        grp_passed_xlsx = os.path.join(output_dir, 'Grouped_Policy_Passed.xlsx')
        if grouped_passed_rows:
            grouped_passed_rows.sort(key=get_group_key)
            df_grouped_passed = pd.DataFrame(grouped_passed_rows)
            df_grouped_passed.to_csv(grp_passed_csv, index=False)
            df_grouped_passed.to_excel(grp_passed_xlsx, index=False)
            print(f"Saved grouped passed records ({len(df_grouped_passed)} rows) to {grp_passed_csv} and {grp_passed_xlsx}")
        else:
            if os.path.exists(grp_passed_csv): os.remove(grp_passed_csv)
            if os.path.exists(grp_passed_xlsx): os.remove(grp_passed_xlsx)
            print("No policy groups fully passed. Cleared Grouped_Policy_Passed files.")

        grp_failed_csv = os.path.join(output_dir, 'Grouped_Policy_Failed.csv')
        grp_failed_xlsx = os.path.join(output_dir, 'Grouped_Policy_Failed.xlsx')
        if grouped_failed_rows:
            grouped_failed_rows.sort(key=get_group_key)
            df_grouped_failed = pd.DataFrame(grouped_failed_rows)
            df_grouped_failed.to_csv(grp_failed_csv, index=False)
            df_grouped_failed.to_excel(grp_failed_xlsx, index=False)
            print(f"Saved grouped failed records ({len(df_grouped_failed)} rows) to {grp_failed_csv} and {grp_failed_xlsx}")
        else:
            if os.path.exists(grp_failed_csv): os.remove(grp_failed_csv)
            if os.path.exists(grp_failed_xlsx): os.remove(grp_failed_xlsx)
            print("No policy groups failed. Cleared Grouped_Policy_Failed files.")

        if single_grouped_file:
            grp_combined_csv = os.path.join(output_dir, 'Grouped_Policy_Combined.csv')
            grp_combined_xlsx = os.path.join(output_dir, 'Grouped_Policy_Combined.xlsx')
            
            combined_rows = []
            for r in grouped_passed_rows:
                r_copy = dict(r)
                r_copy['Group_Status'] = 'Passed'
                combined_rows.append(r_copy)
            for r in grouped_failed_rows:
                r_copy = dict(r)
                r_copy['Group_Status'] = 'Failed'
                combined_rows.append(r_copy)

            if combined_rows:
                combined_rows.sort(key=get_group_key)
                df_combined = pd.DataFrame(combined_rows)
                cols = ['Group_Status'] + [c for c in df_combined.columns if c != 'Group_Status']
                df_combined = df_combined[cols]
                df_combined.to_csv(grp_combined_csv, index=False)
                df_combined.to_excel(grp_combined_xlsx, index=False)
                print(f"Saved combined grouped records ({len(df_combined)} rows) to {grp_combined_csv} and {grp_combined_xlsx}")
            else:
                if os.path.exists(grp_combined_csv): os.remove(grp_combined_csv)
                if os.path.exists(grp_combined_xlsx): os.remove(grp_combined_xlsx)

    print("Done!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Cross-validate an APL statement against a forecast commission_detail_report.csv."
    )
    parser.add_argument("--apl", help="Path to the APL statement CSV or Excel file")
    parser.add_argument(
        "--parallel",
        help="Path to the forecast commission_detail_report.csv (ICM payables)",
    )
    parser.add_argument("--tolerance", type=float, default=0.01, help="Amount match tolerance (default 0.01)")
    parser.add_argument(
        "--mode",
        choices=list(VALID_MODES),
        help="Primary driving side: apl (APL -> ICM) or icm (ICM -> APL). Prompted when omitted.",
    )
    parser.add_argument(
        "--single-grouped-file",
        action="store_true",
        help="Combine grouped policy passed and failed records into a single file (Grouped_Policy_Combined.csv/xlsx)",
    )
    args = parser.parse_args()

    mode = normalize_mode(args.mode) if args.mode else None
    single_grouped = args.single_grouped_file

    if args.apl and args.parallel:
        apl_file = args.apl.strip().strip('\'"')
        parallel_file = args.parallel.strip().strip('\'"')
        tolerance = args.tolerance
        if mode is None:
            mode = prompt_for_mode()
    else:
        print("Welcome to the Policy Commission Validation Script.")
        apl_file = input("Enter path to APL Statement CSV: ").strip().strip('\'"')
        parallel_file = input(
            "Enter path to forecast commission_detail_report.csv: "
        ).strip().strip('\'"')
        tolerance = 0.01
        use_tol = input("Do you want to set a custom amount tolerance? (y/n) [default: n]: ").strip().lower()
        if use_tol in ['y', 'yes']:
            tol_input = input("Enter amount tolerance (e.g. 0.05, 1.5): ").strip()
            try:
                tolerance = float(tol_input)
            except ValueError:
                print("Invalid input, defaulting to 0.01")
                tolerance = 0.01

        use_single = input("Do you want to combine grouped policy passed and failed into a single file? (y/n) [default: n]: ").strip().lower()
        if use_single in ['y', 'yes']:
            single_grouped = True

        if mode is None:
            mode = prompt_for_mode()

    if mode not in VALID_MODES:
        print(f"Error: Unsupported mode '{mode}'. Expected one of: {', '.join(VALID_MODES)}")
        exit(1)

    if not (os.path.exists(apl_file) and os.path.exists(parallel_file)):
        print("Error: One or more input files do not exist. Please check the paths and try again.")
        exit(1)
        
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
    validate_data(apl_file, parallel_file, out_csv, out_xlsx, tolerance, mode=mode, single_grouped_file=single_grouped)
    print(f"\nValidation complete. Outputs saved in {output_folder}")
