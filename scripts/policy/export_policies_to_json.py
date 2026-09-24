#!/usr/bin/env python3
import os
import sys
import json
import argparse
import pandas as pd

def extract_policy_data(json_data):
    """
    Parses json data (either list of dicts or dict with 'data' list)
    and extracts list of dicts with 'Policy No', 'Carrier', and 'Product'.
    """
    if isinstance(json_data, dict):
        if "data" in json_data and isinstance(json_data["data"], list):
            items = json_data["data"]
        else:
            items = [json_data]
    elif isinstance(json_data, list):
        items = json_data
    else:
        items = []

    records = []
    for item in items:
        if not isinstance(item, dict):
            continue

        # Extract Policy Number
        policy_no = item.get("policy_number") or item.get("policy_no")
        if not policy_no and "data" in item and isinstance(item["data"], dict):
            nested_data = item["data"]
            policy_no = (
                nested_data.get("Policy No")
                or nested_data.get("Processed Policy Number")
                or nested_data.get("policy_number")
                or nested_data.get("policy_no")
            )

        # Extract Carrier
        carrier = item.get("carrier_name") or item.get("carrier")
        if not carrier and "data" in item and isinstance(item["data"], dict):
            nested_data = item["data"]
            carrier = (
                nested_data.get("Carrier")
                or nested_data.get("carrier_name")
                or nested_data.get("carrier")
            )

        # Extract Product
        product = item.get("product_name") or item.get("product")
        if not product and "data" in item and isinstance(item["data"], dict):
            nested_data = item["data"]
            product = (
                nested_data.get("Product")
                or nested_data.get("product_name")
                or nested_data.get("product")
            )

        if policy_no or carrier or product:
            records.append({
                "Policy No": str(policy_no).strip() if policy_no else "",
                "Carrier": str(carrier).strip() if carrier else "",
                "Product": str(product).strip() if product else ""
            })

    return pd.DataFrame(records)

def main():
    parser = argparse.ArgumentParser(
        description="Extract Policy No, Carrier, and Product from JSON to Excel, supporting append and deduplication."
    )
    parser.add_argument(
        "-j", "--json-input",
        type=str,
        required=True,
        help="Path to the JSON file containing policy data, or a raw JSON string."
    )
    parser.add_argument(
        "-o", "--output-excel",
        type=str,
        required=True,
        help="Path to the output Excel file."
    )
    args = parser.parse_args()

    # Ensure output has a valid Excel extension
    _, ext = os.path.splitext(args.output_excel.lower())
    if ext not in (".xlsx", ".xls"):
        args.output_excel += ".xlsx"

    # Try to load JSON from file first, then try raw string
    try:
        if os.path.exists(args.json_input):
            with open(args.json_input, "r") as f:
                json_data = json.load(f)
        else:
            json_data = json.loads(args.json_input)
    except Exception as e:
        print(f"❌ Error parsing JSON input: {e}", file=sys.stderr)
        sys.exit(1)

    # Extract records
    new_df = extract_policy_data(json_data)
    if new_df.empty:
        print("⚠️ No valid policy numbers, carriers, or products found in the JSON input.")
        sys.exit(0)

    print(f"📋 Extracted {len(new_df)} records from JSON.")

    # Load existing Excel file if it exists
    if os.path.exists(args.output_excel):
        try:
            existing_df = pd.read_excel(args.output_excel)
            # Standardize columns to avoid duplicate keys during concat
            if "Policy No" not in existing_df.columns:
                existing_df["Policy No"] = ""
            if "Carrier" not in existing_df.columns:
                existing_df["Carrier"] = ""
            if "Product" not in existing_df.columns:
                existing_df["Product"] = ""
            existing_df = existing_df[["Policy No", "Carrier", "Product"]].astype(str)
            
            combined_df = pd.concat([existing_df, new_df.astype(str)], ignore_index=True)
            action = "Appended to and deduplicated"
        except Exception as e:
            print(f"⚠️ Error reading existing Excel file, overwriting instead: {e}")
            combined_df = new_df
            action = "Created new"
    else:
        combined_df = new_df
        action = "Created new"

    # Normalize values for deduplication
    combined_df["Policy No"] = combined_df["Policy No"].str.strip()
    combined_df["Carrier"] = combined_df["Carrier"].str.strip()
    combined_df["Product"] = combined_df["Product"].str.strip()
    
    # Drop exact duplicates and duplicate combinations
    combined_df = combined_df.drop_duplicates(subset=["Policy No", "Carrier", "Product"], keep="first")
    
    # Remove rows where all fields are empty
    combined_df = combined_df[~((combined_df["Policy No"] == "") & (combined_df["Carrier"] == "") & (combined_df["Product"] == ""))]

    # Save to Excel
    try:
        combined_df.to_excel(args.output_excel, index=False)
        print(f"✅ {action} Excel file: '{args.output_excel}' | Total unique records: {len(combined_df)}")
    except Exception as e:
        print(f"❌ Error writing to Excel file: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
