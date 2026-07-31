#!/usr/bin/env python3
import os
import shutil
import tempfile
import unittest
import pandas as pd
from enrich_policy_commission_import import (
    process_single_split_file,
    propagate_carrier_by_policy,
)


class TestEnrichCommission(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_carrier_commission_percentage_mode(self):
        # Case A: Rates > 0
        base_df = pd.DataFrame({
            "Policy No": ["CLI7115324", "CLI7115324", "CLI7115324"],
            "Product": ["Cancer and Heart", "Cancer and Heart", "Cancer and Heart"],
            "Sequence": [1, 2, 3],
            "From Month": ["1", "1", "1"],
            "Agent Name": ["MATTHEW DAUGHERTY", "GREG TEIPEL", "MAIN LINE BENEFITS"],
            "Rate": ["20.0 %", "25.0 %", "30.0 %"],
            "Amount": ["$0.00", "$0.00", "$0.00"]
        })
        base_file = os.path.join(self.temp_dir, "base_comm_rate.csv")
        base_df.to_csv(base_file, index=False)

        master_df = pd.DataFrame({
            "Policy No": ["CLI7115324"],
            "Product": ["Cancer and Heart"],
            "Agent Name": ["MATTHEW DAUGHERTY"],
            "Carrier": ["Aetna"],
            "LOB": ["Ancillary"],
            "Agent Level": ["LVL5"]
        })

        output_df, _ = process_single_split_file(
            base_file_path=base_file,
            master_df=master_df,
        )

        self.assertIn("Carrier Commission", output_df.columns)
        self.assertIn("Payout Method", output_df.columns)
        self.assertEqual(output_df["Carrier Commission"].tolist(), ["30.00 %", "30.00 %", "30.00 %"])
        self.assertEqual(output_df["Payout Method"].tolist(), ["PERCENTAGE", "PERCENTAGE", "PERCENTAGE"])
        self.assertEqual(output_df["Rate"].tolist(), ["66.66 %", "16.66 %", "16.68 %"])

    def test_carrier_commission_fixed_fee_mode(self):
        # Case B: Rates = 0, Amounts > 0
        base_df = pd.DataFrame({
            "Policy No": ["CLI7115324", "CLI7115324", "CLI7115324"],
            "Product": ["Cancer and Heart", "Cancer and Heart", "Cancer and Heart"],
            "Sequence": [1, 2, 3],
            "From Month": ["1", "1", "1"],
            "Agent Name": ["MATTHEW DAUGHERTY", "GREG TEIPEL", "MAIN LINE BENEFITS"],
            "Rate": ["0.00 %", "0.00 %", "0.00 %"],
            "Amount": ["$100.00", "$120.00", "$150.00"]
        })
        base_file = os.path.join(self.temp_dir, "base_comm_amt.csv")
        base_df.to_csv(base_file, index=False)

        master_df = pd.DataFrame({
            "Policy No": ["CLI7115324"],
            "Product": ["Cancer and Heart"],
            "Agent Name": ["MATTHEW DAUGHERTY"],
            "Carrier": ["Aetna"],
            "LOB": ["Ancillary"],
            "Agent Level": ["LVL5"]
        })

        output_df, _ = process_single_split_file(
            base_file_path=base_file,
            master_df=master_df,
        )

        self.assertIn("Carrier Commission", output_df.columns)
        self.assertIn("Payout Method", output_df.columns)
        self.assertEqual(output_df["Carrier Commission"].tolist(), ["$150.00", "$150.00", "$150.00"])
        self.assertEqual(output_df["Payout Method"].tolist(), ["FIXED FEE", "FIXED FEE", "FIXED FEE"])
        self.assertEqual(output_df["Amount"].tolist(), ["$100.00", "$20.00", "$30.00"])
        self.assertEqual(output_df["Rate"].tolist(), ["66.66 %", "13.33 %", "20.01 %"])


if __name__ == "__main__":
    unittest.main()
