"""Write passed, failed, and extra policy files plus agent totals and a dashboard.

This script does not change the files written by validate_apl_icm.py.
"""

import argparse
import json
import os
import shutil
import sys

import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from validate_apl_icm import (  # noqa: E402
    APLColumns,
    ParallelColumns,
    VALID_MODES,
    _agent_group_key,
    _money,
    _preferred_label,
    align_apl_columns,
    apl_line_amount,
    check_amount_match,
    clean_amount,
    clean_policy_number,
    collapse_one_character_suffixes,
    extract_agent_id,
    load_file,
    match_policy_base,
    mode_title,
    normalize_mode,
    normalize_name,
    output_prefix,
    policy_base_number,
    prompt_for_mode,
)

PIE_LIMIT = 8


def display_policy_number(val):
    """Policy number with quote marks removed. OVR and letter suffixes stay."""
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return ""
    text = str(val).strip().replace("'", "").replace('"', "").strip()
    if text.lower() == "nan":
        return ""
    return text


def cell_text(row, column):
    if row is None or column not in row.index:
        return ""
    if pd.isna(row[column]):
        return ""
    text = str(row[column]).strip()
    return "" if text.lower() == "nan" else text


def align_icm_columns(df):
    normalized = {str(column).strip().lower(): column for column in df.columns}
    aliases = {
        ParallelColumns.POLICY_NUMBER: ["policy_number", "policy number", "policy no", "policy no.", "policynumber"],
        ParallelColumns.AGENT_NAME: ["agent name", "agent_name", "writing agent name"],
        ParallelColumns.AMOUNT: ["commission_amount", "commission amount", "amount"],
        ParallelColumns.PRODUCT: ["product_name", "product name", "product"],
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
    missing = [name for name in (ParallelColumns.POLICY_NUMBER, ParallelColumns.AGENT_NAME, ParallelColumns.AMOUNT) if name not in df.columns]
    if missing:
        raise KeyError(
            "ICM payables are missing required column(s): "
            + ", ".join(missing)
            + ". Found columns: "
            + ", ".join(str(column) for column in df.columns)
        )
    return df


def build_comparison_rows(apl_path, icm_path, amount_tolerance):
    """Pair APL and ICM lines, then keep leftovers from both sides."""
    df_apl = align_apl_columns(load_file(apl_path))
    df_icm = align_icm_columns(load_file(icm_path))
    df_apl["Cleaned Policy Number"] = df_apl[APLColumns.POLICY_NUMBER].apply(clean_policy_number)

    icm_pool = {}
    for idx, row in df_icm.iterrows():
        policy = clean_policy_number(row[ParallelColumns.POLICY_NUMBER])
        icm_pool.setdefault(policy, []).append({"row": row, "consumed": False, "idx": idx})

    apl_items = []
    for idx, row in df_apl.iterrows():
        cleaned = str(row["Cleaned Policy Number"])
        target, is_partial = match_policy_base(cleaned, icm_pool)
        apl_items.append({
            "row": row,
            "idx": idx,
            "target_pol": target,
            "is_partial": is_partial,
            "matched_par": None,
        })

    def bind_match(apl_item, icm_item):
        apl_item["matched_par"] = icm_item
        icm_item["consumed"] = True

    for target_pol, icm_list in icm_pool.items():
        target_apl = [item for item in apl_items if item["target_pol"] == target_pol]
        if not target_apl:
            continue

        for apl_item in target_apl:
            if apl_item["matched_par"]:
                continue
            apl_agent = cell_text(apl_item["row"], APLColumns.AGENT_NAME)
            for icm_item in icm_list:
                if icm_item["consumed"]:
                    continue
                icm_agent = cell_text(icm_item["row"], ParallelColumns.AGENT_NAME)
                icm_amount = clean_amount(cell_text(icm_item["row"], ParallelColumns.AMOUNT))
                agent_ok = normalize_name(apl_agent) == normalize_name(icm_agent)
                amount_col, _ = check_amount_match(apl_item["row"], icm_amount, amount_tolerance)
                if agent_ok and amount_col != "None":
                    bind_match(apl_item, icm_item)
                    break

        for apl_item in target_apl:
            if apl_item["matched_par"]:
                continue
            apl_agent = cell_text(apl_item["row"], APLColumns.AGENT_NAME)
            _, _, apl_amount = apl_line_amount(apl_item["row"])
            best_item = None
            best_diff = None
            for icm_item in icm_list:
                if icm_item["consumed"]:
                    continue
                icm_agent = cell_text(icm_item["row"], ParallelColumns.AGENT_NAME)
                if normalize_name(apl_agent) != normalize_name(icm_agent):
                    continue
                diff = abs(apl_amount - clean_amount(cell_text(icm_item["row"], ParallelColumns.AMOUNT)))
                if best_item is None or diff < best_diff:
                    best_item = icm_item
                    best_diff = diff
            if best_item is not None:
                bind_match(apl_item, best_item)

        for apl_item in target_apl:
            if apl_item["matched_par"]:
                continue
            for icm_item in icm_list:
                if icm_item["consumed"]:
                    continue
                icm_amount = clean_amount(cell_text(icm_item["row"], ParallelColumns.AMOUNT))
                amount_col, _ = check_amount_match(apl_item["row"], icm_amount, amount_tolerance)
                if amount_col != "None":
                    bind_match(apl_item, icm_item)
                    break

        for apl_item in target_apl:
            if apl_item["matched_par"]:
                continue
            for icm_item in icm_list:
                if icm_item["consumed"]:
                    continue
                bind_match(apl_item, icm_item)
                break

    rows = []
    order = 0
    for apl_item in apl_items:
        match = apl_item["matched_par"]
        rows.append(comparison_row(apl_item["row"], match["row"] if match else None, amount_tolerance, order))
        order += 1
    for icm_list in icm_pool.values():
        for icm_item in icm_list:
            if icm_item["consumed"]:
                continue
            icm_amount = clean_amount(cell_text(icm_item["row"], ParallelColumns.AMOUNT))
            if abs(icm_amount) <= amount_tolerance:
                continue
            rows.append(comparison_row(None, icm_item["row"], amount_tolerance, order))
            order += 1
    assign_grouped_policies(rows)
    return rows


def comparison_row(apl_row, icm_row, amount_tolerance, order):
    paired = apl_row is not None and icm_row is not None
    apl_policy = display_policy_number(apl_row[APLColumns.POLICY_NUMBER]) if apl_row is not None else ""
    icm_policy = display_policy_number(icm_row[ParallelColumns.POLICY_NUMBER]) if icm_row is not None else ""
    apl_agent = cell_text(apl_row, APLColumns.AGENT_NAME) if apl_row is not None else ""
    icm_agent = cell_text(icm_row, ParallelColumns.AGENT_NAME) if icm_row is not None else ""
    apl_amount = round(apl_line_amount(apl_row)[2], 2) if apl_row is not None else 0.0
    icm_amount = round(clean_amount(cell_text(icm_row, ParallelColumns.AMOUNT)), 2) if icm_row is not None else 0.0
    agent_match = paired and normalize_name(apl_agent) == normalize_name(icm_agent)
    if paired:
        amount_col, _ = check_amount_match(apl_row, icm_amount, amount_tolerance)
        amount_match = amount_col != "None"
    else:
        amount_match = False
    if apl_row is not None and icm_row is None:
        side = "APL"
    elif icm_row is not None and apl_row is None:
        side = "ICM"
    else:
        side = "BOTH"
    return {
        "order": order,
        "apl_policy": apl_policy,
        "icm_policy": icm_policy,
        "apl_agent": apl_agent,
        "icm_agent": icm_agent,
        "apl_agent_id": extract_agent_id(apl_row),
        "icm_agent_id": extract_agent_id(icm_row),
        "apl_amount": apl_amount,
        "icm_amount": icm_amount,
        "paired": paired,
        "agent_match": agent_match,
        "amount_match": amount_match,
        "side": side,
        "base": policy_base_number(icm_policy) or policy_base_number(apl_policy),
    }


def assign_grouped_policies(rows):
    aliases = collapse_one_character_suffixes([row["base"] for row in rows])
    for row in rows:
        row["grouped"] = aliases.get(row["base"], row["base"])


def classify_policies(rows):
    groups = {}
    for row in rows:
        groups.setdefault(row["grouped"], []).append(row)

    passed_rows = []
    failed_rows = []
    extra_rows = []
    for grouped, group_rows in groups.items():
        status = policy_status(group_rows)
        ordered = sorted(group_rows, key=lambda row: row["order"])
        if status == "passed":
            passed_rows.extend(ordered)
        elif status == "extra":
            extra_rows.extend(ordered)
        else:
            for row in ordered:
                row["reason"] = row_reason(row, ordered)
            failed_rows.extend(ordered)
    return passed_rows, failed_rows, extra_rows


def policy_status(group_rows):
    if all(not row["paired"] for row in group_rows):
        return "extra"
    if all(row["paired"] and row["agent_match"] and row["amount_match"] for row in group_rows):
        return "passed"
    return "failed"


def row_reason(row, group_rows):
    if not row["paired"]:
        return "No ICM line" if row["side"] == "APL" else "No APL line"
    if not row["agent_match"]:
        return "Agent name does not match"
    if not row["amount_match"]:
        return "Amount outside tolerance"
    if any(not item["paired"] for item in group_rows):
        return "Policy has a one-sided line"
    if any(item["paired"] and not item["agent_match"] for item in group_rows):
        return "Policy has an agent mismatch"
    return "Policy has an amount mismatch"


def line_frame(rows, include_reason):
    records = []
    for row in rows:
        record = {
            "Grouped Policy": row["grouped"],
            "APL Policy Number": row["apl_policy"],
            "ICM Policy Number": row["icm_policy"],
            "APL Agent Name": row["apl_agent"],
            "ICM Agent Name": row["icm_agent"],
            "APL Amount": row["apl_amount"],
            "ICM Amount": row["icm_amount"],
            "Difference": round(row["icm_amount"] - row["apl_amount"], 2),
        }
        if include_reason:
            record["Reason"] = row["reason"]
        records.append(record)
    columns = [
        "Grouped Policy", "APL Policy Number", "ICM Policy Number",
        "APL Agent Name", "ICM Agent Name", "APL Amount", "ICM Amount", "Difference",
    ]
    if include_reason:
        columns.append("Reason")
    return pd.DataFrame(records, columns=columns)


def extra_frame(rows):
    records = []
    for row in rows:
        records.append({
            "Grouped Policy": row["grouped"],
            "Side": row["side"],
            "APL Policy Number": row["apl_policy"],
            "ICM Policy Number": row["icm_policy"],
            "Agent Name": row["apl_agent"] or row["icm_agent"],
            "APL Amount": row["apl_amount"],
            "ICM Amount": row["icm_amount"],
        })
    return pd.DataFrame(records, columns=[
        "Grouped Policy", "Side", "APL Policy Number", "ICM Policy Number",
        "Agent Name", "APL Amount", "ICM Amount",
    ])


def remember(counts, label):
    if label:
        counts[label] = counts.get(label, 0) + 1


def agent_tables(rows, failed_rows):
    agents = {}
    mismatch_policies = {}

    def bucket(key):
        return agents.setdefault(key, {
            "apl_names": {},
            "icm_names": {},
            "ids": {},
            "apl_amount": 0.0,
            "icm_amount": 0.0,
            "apl_amount_excluding_extra": 0.0,
        })

    mismatched_policies = set()
    for row in failed_rows:
        if row["paired"] and (not row["agent_match"] or not row["amount_match"]):
            mismatched_policies.add(row["grouped"])

    for row in rows:
        key = _agent_group_key(row["apl_agent"], row["icm_agent"])
        agent = bucket(key)
        remember(agent["apl_names"], row["apl_agent"])
        remember(agent["icm_names"], row["icm_agent"])
        agent_id = row["icm_agent_id"] or row["apl_agent_id"]
        remember(agent["ids"], agent_id)
        agent["apl_amount"] += row["apl_amount"]
        agent["icm_amount"] += row["icm_amount"]
        if row["side"] != "APL":
            agent["apl_amount_excluding_extra"] += row["apl_amount"]

        if row["grouped"] in mismatched_policies and row["paired"] and (not row["agent_match"] or not row["amount_match"]):
            entry = mismatch_policies.setdefault((key, row["grouped"]), {
                "apl_names": {},
                "icm_names": {},
                "ids": {},
                "apl_amount": 0.0,
                "icm_amount": 0.0,
            })
            remember(entry["apl_names"], row["apl_agent"])
            remember(entry["icm_names"], row["icm_agent"])
            remember(entry["ids"], agent_id)
            entry["apl_amount"] += row["apl_amount"]
            entry["icm_amount"] += row["icm_amount"]

    total_rows = []
    excluded_rows = []
    for key, agent in agents.items():
        name = _preferred_label(agent["icm_names"]) or _preferred_label(agent["apl_names"]) or "(No agent name)"
        agent_id = _preferred_label(agent["ids"])
        total_rows.append({
            "Agent Name": name,
            "Agent ID": agent_id,
            "APL Amount": round(agent["apl_amount"], 2),
            "ICM Amount": round(agent["icm_amount"], 2),
            "Difference": round(agent["icm_amount"] - agent["apl_amount"], 2),
        })
        excluded_rows.append({
            "Agent Name": name,
            "Agent ID": agent_id,
            "APL Amount": round(agent["apl_amount_excluding_extra"], 2),
            "ICM Amount": round(agent["icm_amount"], 2),
            "Difference": round(agent["icm_amount"] - agent["apl_amount_excluding_extra"], 2),
        })

    by_agent = {}
    for (key, policy_number), entry in mismatch_policies.items():
        agent = by_agent.setdefault(key, {
            "apl_names": {},
            "icm_names": {},
            "ids": {},
            "apl_amount": 0.0,
            "icm_amount": 0.0,
            "policies": [],
        })
        agent["apl_names"].update(entry["apl_names"])
        agent["icm_names"].update(entry["icm_names"])
        agent["ids"].update(entry["ids"])
        agent["apl_amount"] += entry["apl_amount"]
        agent["icm_amount"] += entry["icm_amount"]
        agent["policies"].append(policy_number)

    mismatch_rows = []
    for agent in by_agent.values():
        policies = sorted(set(agent["policies"]))
        apl_amount = round(agent["apl_amount"], 2)
        icm_amount = round(agent["icm_amount"], 2)
        mismatch_rows.append({
            "ICM Agent Name": _preferred_label(agent["icm_names"]),
            "APL Agent Name": _preferred_label(agent["apl_names"]),
            "Agent ID": _preferred_label(agent["ids"]),
            "APL Amount": apl_amount,
            "ICM Amount": icm_amount,
            "Amount Difference": round(icm_amount - apl_amount, 2),
            "Grouped Policy Count": len(policies),
            "Policy Numbers": ", ".join(policies),
        })

    def by_difference(record):
        difference = record.get("Difference", record.get("Amount Difference", 0.0))
        name = record.get("Agent Name") or record.get("ICM Agent Name") or ""
        return (-abs(difference), name.lower())

    total_rows.sort(key=by_difference)
    excluded_rows.sort(key=by_difference)
    mismatch_rows.sort(key=by_difference)
    return (
        pd.DataFrame(total_rows, columns=["Agent Name", "Agent ID", "APL Amount", "ICM Amount", "Difference"]),
        pd.DataFrame(excluded_rows, columns=["Agent Name", "Agent ID", "APL Amount", "ICM Amount", "Difference"]),
        pd.DataFrame(mismatch_rows, columns=[
            "ICM Agent Name", "APL Agent Name", "Agent ID", "APL Amount", "ICM Amount",
            "Amount Difference", "Grouped Policy Count", "Policy Numbers",
        ]),
    )


def write_frame(frame, path_no_ext):
    frame.to_csv(f"{path_no_ext}.csv", index=False, float_format="%.2f")
    frame.to_excel(f"{path_no_ext}.xlsx", index=False)


def policy_count(rows):
    return len({row["grouped"] for row in rows})


def write_summary(path, mode, apl_filename, icm_filename, amount_tolerance, passed_rows, failed_rows, extra_rows, totals, excluded, mismatches):
    passed_count = policy_count(passed_rows)
    failed_count = policy_count(failed_rows)
    extra_count = policy_count(extra_rows)
    apl_total = float(totals["APL Amount"].sum()) if not totals.empty else 0.0
    icm_total = float(totals["ICM Amount"].sum()) if not totals.empty else 0.0
    excluded_apl = float(excluded["APL Amount"].sum()) if not excluded.empty else 0.0
    excluded_icm = float(excluded["ICM Amount"].sum()) if not excluded.empty else 0.0

    def agent_lines(frame, name_column, difference_column):
        if frame.empty:
            return "| — | — | — | — |"
        lines = []
        for _, agent in frame.head(20).iterrows():
            lines.append(
                f"| {agent[name_column]} | {agent['Agent ID']} | {_money(agent['APL Amount'])} | "
                f"{_money(agent['ICM Amount'])} | {_money(agent[difference_column])} |"
            )
        return "\n".join(lines)

    mismatch_lines = []
    if mismatches.empty:
        mismatch_lines.append("| — | — | — | — | — |")
    else:
        for _, agent in mismatches.head(20).iterrows():
            mismatch_lines.append(
                f"| {agent['ICM Agent Name']} | {agent['APL Agent Name']} | {_money(agent['Amount Difference'])} | "
                f"{int(agent['Grouped Policy Count'])} | {agent['Policy Numbers']} |"
            )

    title = mode_title(mode)
    content = f"""# {title} policy group report

| Input | File |
| --- | --- |
| APL statement | `{apl_filename}` |
| ICM payables | `{icm_filename}` |
| Amount tolerance | {amount_tolerance} |
| Comparison | {title} |

| Result | Policies |
| --- | ---: |
| Passed | {passed_count} |
| Failed | {failed_count} |
| Extra | {extra_count} |

All-lines APL {_money(apl_total)}, ICM {_money(icm_total)}, difference {_money(round(icm_total - apl_total, 2))}.

APL excluding extra APL lines {_money(excluded_apl)}, ICM {_money(excluded_icm)}, difference {_money(round(excluded_icm - excluded_apl, 2))}.

## Agent totals

| Agent Name | Agent ID | APL Amount | ICM Amount | Difference |
| --- | --- | ---: | ---: | ---: |
{agent_lines(totals, "Agent Name", "Difference")}

## Agent totals excluding extra APL lines

| Agent Name | Agent ID | APL Amount | ICM Amount | Difference |
| --- | --- | ---: | ---: | ---: |
{agent_lines(excluded, "Agent Name", "Difference")}

## Agent amount mismatches

Policies whose paired agent or amount does not match. A policy that failed only because of a one-sided line is not listed.

| ICM Agent Name | APL Agent Name | Amount Difference | Grouped Policy Count | Policy Numbers |
| --- | --- | ---: | ---: | --- |
{chr(10).join(mismatch_lines)}
"""
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(content)


def pie_slices(pairs, limit=PIE_LIMIT):
    usable = [(label, float(value)) for label, value in pairs if float(value) > 0]
    usable.sort(key=lambda item: (-item[1], item[0].lower()))
    head = usable[:limit]
    rest = round(sum(value for _, value in usable[limit:]), 2)
    slices = [{"label": label, "value": round(value, 2)} for label, value in head]
    if rest > 0:
        slices.append({"label": "Other", "value": rest})
    return slices


def dashboard_payload(mode, amount_tolerance, passed_rows, failed_rows, extra_rows, totals, excluded, mismatches):
    apl_total = float(totals["APL Amount"].sum()) if not totals.empty else 0.0
    icm_total = float(totals["ICM Amount"].sum()) if not totals.empty else 0.0
    excluded_apl = float(excluded["APL Amount"].sum()) if not excluded.empty else 0.0
    excluded_icm = float(excluded["ICM Amount"].sum()) if not excluded.empty else 0.0

    def records(frame):
        if frame.empty:
            return []
        return json.loads(frame.to_json(orient="records"))

    return {
        "title": mode_title(mode),
        "tolerance": amount_tolerance,
        "policies": {
            "passed": policy_count(passed_rows),
            "failed": policy_count(failed_rows),
            "extra": policy_count(extra_rows),
        },
        "totals": {
            "apl": round(apl_total, 2),
            "icm": round(icm_total, 2),
            "difference": round(icm_total - apl_total, 2),
            "aplExcludingExtra": round(excluded_apl, 2),
            "icmExcludingExtra": round(excluded_icm, 2),
            "differenceExcludingExtra": round(excluded_icm - excluded_apl, 2),
        },
        "pies": {
            "policies": [
                {"label": "Passed", "value": policy_count(passed_rows)},
                {"label": "Failed", "value": policy_count(failed_rows)},
                {"label": "Extra", "value": policy_count(extra_rows)},
            ],
            "apl": pie_slices(zip(totals["Agent Name"], totals["APL Amount"]) if not totals.empty else []),
            "icm": pie_slices(zip(totals["Agent Name"], totals["ICM Amount"]) if not totals.empty else []),
            "difference": pie_slices(zip(totals["Agent Name"], totals["Difference"].abs()) if not totals.empty else []),
        },
        "agents": records(totals),
        "agentsExcludingExtraApl": records(excluded),
        "mismatches": records(mismatches),
    }


def write_dashboard(path, payload):
    data = json.dumps(payload).replace("<", "\\u003c")
    html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Policy group dashboard</title>
<style>
body { font-family: Georgia, serif; margin: 24px; color: #1c1c1c; background: #f7f6f3; }
h1, h2 { font-weight: normal; }
.cards { display: flex; flex-wrap: wrap; gap: 12px; margin: 16px 0 28px; }
.card { background: #fff; border: 1px solid #ddd; padding: 12px 16px; min-width: 160px; }
.card span { display: block; font-size: 12px; letter-spacing: 0.04em; text-transform: uppercase; color: #555; }
.card strong { font-size: 22px; }
.pies { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 16px; }
.pie { background: #fff; border: 1px solid #ddd; padding: 12px; }
.pie-row { display: flex; gap: 12px; align-items: center; }
ul { list-style: none; padding: 0; margin: 0; font-size: 13px; }
li { margin: 4px 0; }
.swatch { display: inline-block; width: 10px; height: 10px; margin-right: 6px; }
table { width: 100%; border-collapse: collapse; background: #fff; margin: 8px 0 28px; }
th, td { border-bottom: 1px solid #e4e4e4; text-align: left; padding: 6px 8px; font-size: 14px; }
th { font-weight: normal; color: #444; }
input { margin: 8px 0; padding: 6px 8px; width: min(360px, 100%); }
</style>
</head>
<body>
<h1 id="title"></h1>
<p id="tolerance"></p>
<div class="cards" id="cards"></div>
<div class="pies" id="pies"></div>
<h2>Agent totals</h2>
<input id="agent-search" placeholder="Search agents">
<div id="agent-table"></div>
<h2>Agent totals excluding extra APL</h2>
<div id="excluded-table"></div>
<h2>Agent amount mismatches</h2>
<input id="mismatch-search" placeholder="Search mismatches or policy numbers">
<div id="mismatch-table"></div>
<script>
const DATA = __DATA__;
const COLORS = ["#3d5a40", "#588157", "#a3b18a", "#bc6c25", "#606c38", "#283618", "#dda15e", "#6b705c", "#9c6644", "#b7b7a4"];

function money(value) {
  const amount = Number(value || 0);
  const sign = amount < 0 ? "-" : "";
  return sign + "$" + Math.abs(amount).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function card(label, value) {
  const node = document.createElement("div");
  node.className = "card";
  const span = document.createElement("span");
  span.textContent = label;
  const strong = document.createElement("strong");
  strong.textContent = value;
  node.append(span, strong);
  return node;
}

document.getElementById("title").textContent = DATA.title + " policy group dashboard";
document.getElementById("tolerance").textContent = "Amount tolerance " + DATA.tolerance;
const cards = document.getElementById("cards");
cards.append(
  card("Passed policies", DATA.policies.passed),
  card("Failed policies", DATA.policies.failed),
  card("Extra policies", DATA.policies.extra),
  card("APL amount", money(DATA.totals.apl)),
  card("ICM amount", money(DATA.totals.icm)),
  card("Difference", money(DATA.totals.difference)),
  card("APL excluding extra", money(DATA.totals.aplExcludingExtra)),
  card("Difference excluding extra APL", money(DATA.totals.differenceExcludingExtra))
);

function drawPie(canvas, slices) {
  const ctx = canvas.getContext("2d");
  const size = canvas.width;
  const radius = size / 2 - 8;
  const center = size / 2;
  ctx.clearRect(0, 0, size, size);
  const total = slices.reduce((sum, slice) => sum + Number(slice.value || 0), 0);
  if (!total) {
    ctx.beginPath();
    ctx.arc(center, center, radius, 0, Math.PI * 2);
    ctx.strokeStyle = "#ccc";
    ctx.stroke();
    return;
  }
  let angle = -Math.PI / 2;
  slices.forEach((slice, index) => {
    const sweep = (Number(slice.value) / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(center, center);
    ctx.arc(center, center, radius, angle, angle + sweep);
    ctx.closePath();
    ctx.fillStyle = COLORS[index % COLORS.length];
    ctx.fill();
    angle += sweep;
  });
}

function legend(slices) {
  const list = document.createElement("ul");
  slices.forEach((slice, index) => {
    const item = document.createElement("li");
    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.style.background = COLORS[index % COLORS.length];
    item.append(swatch, document.createTextNode(slice.label + " (" + slice.value + ")"));
    list.append(item);
  });
  return list;
}

function pieBlock(heading, slices) {
  const block = document.createElement("section");
  block.className = "pie";
  const title = document.createElement("h2");
  title.textContent = heading;
  const row = document.createElement("div");
  row.className = "pie-row";
  const canvas = document.createElement("canvas");
  canvas.width = 180;
  canvas.height = 180;
  drawPie(canvas, slices);
  row.append(canvas, legend(slices));
  block.append(title, row);
  return block;
}

const pies = document.getElementById("pies");
pies.append(
  pieBlock("Policies", DATA.pies.policies),
  pieBlock("APL amount by agent", DATA.pies.apl),
  pieBlock("ICM amount by agent", DATA.pies.icm),
  pieBlock("Absolute difference by agent", DATA.pies.difference)
);

function table(columns, rows) {
  const node = document.createElement("table");
  const head = document.createElement("tr");
  columns.forEach((column) => {
    const cell = document.createElement("th");
    cell.textContent = column.label;
    head.append(cell);
  });
  const thead = document.createElement("thead");
  thead.append(head);
  const body = document.createElement("tbody");
  rows.forEach((row) => {
    const line = document.createElement("tr");
    columns.forEach((column) => {
      const cell = document.createElement("td");
      const value = row[column.key];
      cell.textContent = column.money ? money(value) : (value ?? "");
      line.append(cell);
    });
    body.append(line);
  });
  node.append(thead, body);
  return node;
}

const agentColumns = [
  {key: "Agent Name", label: "Agent Name"},
  {key: "Agent ID", label: "Agent ID"},
  {key: "APL Amount", label: "APL Amount", money: true},
  {key: "ICM Amount", label: "ICM Amount", money: true},
  {key: "Difference", label: "Difference", money: true}
];
const mismatchColumns = [
  {key: "ICM Agent Name", label: "ICM Agent"},
  {key: "APL Agent Name", label: "APL Agent"},
  {key: "Agent ID", label: "Agent ID"},
  {key: "Amount Difference", label: "Amount Difference", money: true},
  {key: "Grouped Policy Count", label: "Policies"},
  {key: "Policy Numbers", label: "Policy Numbers"}
];

function mount(id, columns, rows) {
  const host = document.getElementById(id);
  host.replaceChildren(table(columns, rows));
}

function matches(row, query) {
  return Object.values(row).join(" ").toLowerCase().includes(query);
}

mount("agent-table", agentColumns, DATA.agents);
mount("excluded-table", agentColumns, DATA.agentsExcludingExtraApl);
mount("mismatch-table", mismatchColumns, DATA.mismatches);

document.getElementById("agent-search").addEventListener("input", (event) => {
  const query = event.target.value.trim().toLowerCase();
  mount("agent-table", agentColumns, DATA.agents.filter((row) => matches(row, query)));
  mount("excluded-table", agentColumns, DATA.agentsExcludingExtraApl.filter((row) => matches(row, query)));
});
document.getElementById("mismatch-search").addEventListener("input", (event) => {
  const query = event.target.value.trim().toLowerCase();
  mount("mismatch-table", mismatchColumns, DATA.mismatches.filter((row) => matches(row, query)));
});
</script>
</body>
</html>
"""
    html = html.replace("__DATA__", data)
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(html)


def report_folders(output_dir):
    folders = {
        "policies": os.path.join(output_dir, "policies"),
        "agents": os.path.join(output_dir, "agents"),
        "summary": os.path.join(output_dir, "summary"),
    }
    for folder in folders.values():
        os.makedirs(folder, exist_ok=True)
    return folders


def write_policy_group_report(apl_path, icm_path, output_dir, amount_tolerance, mode):
    mode = normalize_mode(mode) or "apl"
    if mode not in VALID_MODES:
        raise ValueError(f"Unsupported mode '{mode}'. Expected one of: {', '.join(VALID_MODES)}")
    prefix = output_prefix(mode)
    rows = build_comparison_rows(apl_path, icm_path, amount_tolerance)
    passed_rows, failed_rows, extra_rows = classify_policies(rows)
    totals, excluded, mismatches = agent_tables(rows, failed_rows)
    folders = report_folders(output_dir)

    write_frame(line_frame(passed_rows, include_reason=False), os.path.join(folders["policies"], f"{prefix}_policies_passed"))
    write_frame(line_frame(failed_rows, include_reason=True), os.path.join(folders["policies"], f"{prefix}_policies_failed"))
    write_frame(extra_frame(extra_rows), os.path.join(folders["policies"], f"{prefix}_policies_extra"))
    write_frame(totals, os.path.join(folders["agents"], f"{prefix}_agent_totals"))
    write_frame(excluded, os.path.join(folders["agents"], f"{prefix}_agent_totals_excluding_extra_apl"))
    write_frame(mismatches, os.path.join(folders["agents"], f"{prefix}_agent_amount_mismatches"))

    summary_path = os.path.join(folders["summary"], f"{prefix}_summary.md")
    dashboard_path = os.path.join(folders["summary"], f"{prefix}_dashboard.html")
    write_summary(
        summary_path,
        mode,
        os.path.basename(apl_path),
        os.path.basename(icm_path),
        amount_tolerance,
        passed_rows,
        failed_rows,
        extra_rows,
        totals,
        excluded,
        mismatches,
    )
    write_dashboard(
        dashboard_path,
        dashboard_payload(mode, amount_tolerance, passed_rows, failed_rows, extra_rows, totals, excluded, mismatches),
    )
    print(f"Passed policies: {policy_count(passed_rows)}")
    print(f"Failed policies: {policy_count(failed_rows)}")
    print(f"Extra policies: {policy_count(extra_rows)}")
    print(f"Wrote policy group report to {output_dir}")
    return {
        "passed": passed_rows,
        "failed": failed_rows,
        "extra": extra_rows,
        "totals": totals,
        "excluded": excluded,
        "mismatches": mismatches,
    }


def next_version_folder():
    versions_dir = os.path.join(SCRIPT_DIR, "versions")
    os.makedirs(versions_dir, exist_ok=True)
    existing = [name for name in os.listdir(versions_dir) if name.startswith("v") and os.path.isdir(os.path.join(versions_dir, name))]
    numbers = [int(name[1:]) for name in existing if name[1:].isdigit()]
    next_number = max(numbers) + 1 if numbers else 1
    version_folder = os.path.join(versions_dir, f"v{next_number}")
    input_folder = os.path.join(version_folder, "input")
    output_folder = os.path.join(version_folder, "output")
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    return version_folder, input_folder, output_folder


def main():
    parser = argparse.ArgumentParser(description="Group APL and ICM policies into passed, failed, and extra files.")
    parser.add_argument("--apl", help="Path to the APL statement CSV or Excel file")
    parser.add_argument("--parallel", help="Path to the ICM payables CSV or Excel file")
    parser.add_argument("--tolerance", type=float, default=0.01, help="Amount match tolerance (default 0.01)")
    parser.add_argument("--mode", choices=list(VALID_MODES), help="apl (APL -> ICM) or icm (ICM -> APL). Prompted when omitted.")
    args = parser.parse_args()

    mode = normalize_mode(args.mode) if args.mode else None
    if args.apl and args.parallel:
        apl_file = args.apl.strip().strip("'\"")
        icm_file = args.parallel.strip().strip("'\"")
        tolerance = args.tolerance
        if mode is None:
            mode = prompt_for_mode()
    else:
        print("Policy group report")
        apl_file = input("Enter path to APL Statement CSV: ").strip().strip("'\"")
        icm_file = input("Enter path to ICM payables CSV: ").strip().strip("'\"")
        tolerance = 0.01
        use_tol = input("Do you want to set a custom amount tolerance? (y/n) [default: n]: ").strip().lower()
        if use_tol in ("y", "yes"):
            try:
                tolerance = float(input("Enter amount tolerance (e.g. 0.02): ").strip())
            except ValueError:
                print("Invalid input, defaulting to 0.01")
                tolerance = 0.01
        if mode is None:
            mode = prompt_for_mode()

    if mode not in VALID_MODES:
        print(f"Error: Unsupported mode '{mode}'. Expected one of: {', '.join(VALID_MODES)}")
        sys.exit(1)
    if not (os.path.exists(apl_file) and os.path.exists(icm_file)):
        print("Error: One or more input files do not exist.")
        sys.exit(1)

    version_folder, input_folder, output_folder = next_version_folder()
    print(f"Creating version folder: {version_folder}")
    shutil.copy2(apl_file, os.path.join(input_folder, os.path.basename(apl_file)))
    shutil.copy2(icm_file, os.path.join(input_folder, os.path.basename(icm_file)))
    write_policy_group_report(apl_file, icm_file, output_folder, tolerance, mode)


if __name__ == "__main__":
    main()
