#!/usr/bin/env python3
"""
Unit tests for validate_extracted_commission.py
"""

import os
import sys
import unittest
import pandas as pd
import tempfile
import shutil

# Ensure parent directory is in sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from validate_extracted_commission import (
    parse_amount,
    parse_rate,
    unpivot_product_commission_rates,
    load_product_normalization,
    validate_commission_dataset,
    process_commission_validation,
)


class TestValidateExtractedCommission(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_parse_amount_and_rate(self):
        self.assertEqual(parse_amount("$10.50"), 10.50)
        self.assertEqual(parse_amount("1,250.00"), 1250.00)
        self.assertEqual(parse_amount(None), 0.0)

        self.assertEqual(parse_rate("70.00 %"), 70.00)
        self.assertEqual(parse_rate("15.5%"), 15.5)
        self.assertEqual(parse_rate(None), 0.0)

    def test_unpivot_product_commission_rates(self):
        # Create a mock ProductWise Commission Rate excel file
        mock_data = {
            "Product ID": [101, 102],
            "Product Description": ["Health Shield Plan", "Dental Basic"],
            "Contract": ["LVL1", "AGENCY"],
            "From-To Months": ["1-12", "1-12"],
            "Rate (Percentage)": [0.70, 0.0],
            "Rate (Dollar Amount)": [0.0, 25.0],
            "Column Rate is Fixed": ["N", "Y"],
            "From-To Months.1": ["13-0", ""],
            "Rate (Percentage).1": [0.35, 0.0],
            "Rate (Dollar Amount).1": [0.0, 0.0],
            "Column Rate is Fixed.1": ["N", ""],
        }
        mock_df = pd.DataFrame(mock_data)
        excel_path = os.path.join(self.test_dir, "mock_rates.xlsx")
        mock_df.to_excel(excel_path, index=False)

        rules_df = unpivot_product_commission_rates(excel_path)
        self.assertFalse(rules_df.empty)
        self.assertIn("Health Shield Plan", rules_df["Product_Description"].values)
        self.assertIn("Dental Basic", rules_df["Product_Description"].values)

        # Check unpivoted tier ranges
        health_rules = rules_df[rules_df["Product_Description"] == "Health Shield Plan"]
        self.assertEqual(len(health_rules), 2)
        tier1 = health_rules[health_rules["From_Month"] == 1].iloc[0]
        self.assertAlmostEqual(tier1["Rate_Percentage"], 70.00)
        self.assertEqual(tier1["To_Month"], 12)

        tier2 = health_rules[health_rules["From_Month"] == 13].iloc[0]
        self.assertAlmostEqual(tier2["Rate_Percentage"], 35.00)
        self.assertEqual(tier2["To_Month"], 9999)

    def test_validate_commission_dataset(self):
        # Mock rate rules
        mock_rules = pd.DataFrame([
            {
                "Product_ID": 101,
                "Product_Description": "Health Shield Plan",
                "Contract": "LVL1",
                "From_Month": 1,
                "To_Month": 12,
                "Rate_Percentage": 70.0,
                "Rate_Dollar": 0.0,
                "Is_Fixed": "N",
                "_prod_clean": "health shield plan",
                "_contract_clean": "lvl1",
            },
            {
                "Product_ID": 102,
                "Product_Description": "Dental Basic",
                "Contract": "AGENCY",
                "From_Month": 1,
                "To_Month": 12,
                "Rate_Percentage": 0.0,
                "Rate_Dollar": 25.0,
                "Is_Fixed": "Y",
                "_prod_clean": "dental basic",
                "_contract_clean": "agency",
            },
        ])

        # Mock extracted commission dataset
        mock_comm = pd.DataFrame([
            {
                "Policy No": "POL001",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
            },
            {
                "Policy No": "POL002",
                "Product": "Dental Basic",
                "Agent Level": "AGENCY",
                "From Month": 6,
                "Carrier Commission": "$25.00",
                "Payout Method": "FIXED FEE",
            },
            {
                "Policy No": "POL003",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "50.00 %",
                "Payout Method": "PERCENTAGE",
            },
            {
                "Policy No": "POL004",
                "Product": "Unknown Product",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
            },
        ])

        df_validated, df_errors, summary = validate_commission_dataset(mock_comm, mock_rules)

        self.assertEqual(summary["total_records"], 4)
        self.assertEqual(summary["valid_records"], 2)
        self.assertEqual(summary["value_mismatches"], 1)
        self.assertEqual(summary["unmatched_products"], 1)

        # Check individual row statuses
        df_all = pd.concat([df_validated, df_errors]).sort_index()
        statuses = df_all["Validation_Status"].tolist()
        self.assertEqual(statuses, ["VALID", "VALID", "VALUE_MISMATCH", "UNMATCHED_PRODUCT"])

    def test_validate_commission_dataset_grouping(self):
        # Mock rate rules (Agency contract LVL4 has 75.0% Carrier Commission)
        mock_rules = pd.DataFrame([
            {
                "Product_ID": 201,
                "Product_Description": "BCBS of TX SG",
                "Contract": "LVL4",
                "From_Month": 1,
                "To_Month": 12,
                "Rate_Percentage": 75.0,
                "Rate_Dollar": 0.0,
                "Is_Fixed": "N",
                "_prod_clean": "bcbs of tx sg",
                "_contract_clean": "lvl4",
            }
        ])

        # Mock grouped sequence extracted dataset (Policy 394835 with 2 sequences)
        mock_comm = pd.DataFrame([
            {
                "Policy No": "394835",
                "Product": "BCBS of TX SG",
                "Sequence": 1,
                "Agent Level": "LVL4",
                "From Month": 1,
                "Rate": "80.00 %",
                "Amount": "$0.00",
                "Carrier Commission": "75.00 %",
                "Payout Method": "PERCENTAGE",
            },
            {
                "Policy No": "394835",
                "Product": "BCBS of TX SG",
                "Sequence": 2,
                "Agent Level": "Sales Leader",
                "From Month": 1,
                "Rate": "20.00 %",
                "Amount": "$0.00",
                "Carrier Commission": "75.00 %",
                "Payout Method": "PERCENTAGE",
            },
        ])

        df_validated, df_errors, summary = validate_commission_dataset(mock_comm, mock_rules)

        self.assertEqual(summary["total_records"], 2)
        self.assertEqual(summary["valid_records"], 2)
        self.assertEqual(len(df_errors), 0)
        self.assertEqual(df_validated["Validation_Status"].tolist(), ["VALID", "VALID"])

    def test_carrier_validation(self):
        # Mock rate rules
        mock_rules = pd.DataFrame([
            {
                "Product_ID": 101,
                "Product_Description": "Health Shield Plan",
                "Contract": "LVL1",
                "From_Month": 1,
                "To_Month": 12,
                "Rate_Percentage": 70.0,
                "Rate_Dollar": 0.0,
                "Is_Fixed": "N",
                "_prod_clean": "health shield plan",
                "_contract_clean": "lvl1",
            }
        ])

        # Test case 1: Auto-population (one row has carrier, other doesn't)
        mock_comm_populate = pd.DataFrame([
            {
                "Policy No": "POL001",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "Transamerica",
                "Sequence": 1,
                "Rate": "50.00 %",
            },
            {
                "Policy No": "POL001",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "\\N",
                "Sequence": 2,
                "Rate": "25.00 %",
            },
            {
                "Policy No": "POL001",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "",
                "Sequence": 3,
                "Rate": "25.00 %",
            },
        ])

        df_validated, df_errors, summary = validate_commission_dataset(mock_comm_populate, mock_rules, min_seq_len=3, max_seq_len=4)
        # All rows should have carrier "Transamerica" populated and be VALID
        self.assertEqual(len(df_errors), 0)
        self.assertEqual(df_validated["Carrier"].tolist(), ["Transamerica", "Transamerica", "Transamerica"])

        # Test case 2: Missing carrier completely
        mock_comm_missing = pd.DataFrame([
            {
                "Policy No": "POL002",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "\\N",
                "Sequence": 1,
                "Rate": "50.00 %",
            },
            {
                "Policy No": "POL002",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "",
                "Sequence": 2,
                "Rate": "25.00 %",
            },
            {
                "Policy No": "POL002",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": None,
                "Sequence": 3,
                "Rate": "25.00 %",
            },
        ])

        df_validated, df_errors, summary = validate_commission_dataset(mock_comm_missing, mock_rules, min_seq_len=3, max_seq_len=4)
        self.assertEqual(len(df_validated), 0)
        self.assertEqual(summary["missing_carriers"], 3)
        self.assertEqual(df_errors["Validation_Status"].tolist(), ["MISSING_CARRIER", "MISSING_CARRIER", "MISSING_CARRIER"])

        # Test case 3: Carrier mismatch in the same group
        mock_comm_mismatch = pd.DataFrame([
            {
                "Policy No": "POL003",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "Transamerica",
                "Sequence": 1,
                "Rate": "50.00 %",
            },
            {
                "Policy No": "POL003",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "Aetna",
                "Sequence": 2,
                "Rate": "25.00 %",
            },
            {
                "Policy No": "POL003",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "Transamerica",
                "Sequence": 3,
                "Rate": "25.00 %",
            },
        ])

        df_validated, df_errors, summary = validate_commission_dataset(mock_comm_mismatch, mock_rules, min_seq_len=3, max_seq_len=4)
        self.assertEqual(len(df_validated), 0)
        self.assertEqual(summary["carrier_mismatches"], 3)
        self.assertEqual(df_errors["Validation_Status"].tolist(), ["CARRIER_MISMATCH", "CARRIER_MISMATCH", "CARRIER_MISMATCH"])

    def test_validate_commission_dataset_optional_rules(self):
        # Valid group of 3 sequence rows
        mock_comm = pd.DataFrame([
            {
                "Policy No": "POL001",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "To Month": 12,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "Transamerica",
                "Sequence": 1,
                "Rate": "50.00 %",
            },
            {
                "Policy No": "POL001",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "To Month": 12,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "Transamerica",
                "Sequence": 2,
                "Rate": "25.00 %",
            },
            {
                "Policy No": "POL001",
                "Product": "Health Shield Plan",
                "Agent Level": "LVL1",
                "From Month": 1,
                "To Month": 12,
                "Carrier Commission": "70.00 %",
                "Payout Method": "PERCENTAGE",
                "Carrier": "Transamerica",
                "Sequence": 3,
                "Rate": "25.00 %",
            },
        ])

        df_validated, df_errors, summary = validate_commission_dataset(
            mock_comm, df_rules=None, min_seq_len=3, max_seq_len=4
        )
        self.assertEqual(len(df_validated), 3)
        self.assertEqual(summary["valid_records"], 3)
        self.assertEqual(df_validated["Expected_Base_Commission"].tolist(), ["SKIPPED_RATE_VALIDATION"] * 3)
        self.assertEqual(df_validated["Validation_Status"].tolist(), ["VALID"] * 3)


if __name__ == "__main__":
    unittest.main()
