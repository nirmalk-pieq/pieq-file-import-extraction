#!/usr/bin/env python3
import os
import shutil
import tempfile
import unittest
import pandas as pd
from extract_unique_commission import (
    clean_key,
    extract_policy_keys_from_input,
    process_commission_files,
)


class TestExtractUniqueCommission(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_clean_key(self):
        s = pd.Series(["  p-101  ", "'P-102'", '"P-103"', None])
        cleaned = clean_key(s).tolist()
        self.assertEqual(cleaned, ["P-101", "P-102", "P-103", ""])

    def test_unique_commission_extraction(self):
        # Create mock policy import file
        policy_df = pd.DataFrame({
            "Policy No": ["POL100", "POL200", "POL300"],
            "Customer": ["Alice", "Bob", "Charlie"]
        })
        policy_file = os.path.join(self.temp_dir, "Policy_Import.xlsx")
        policy_df.to_excel(policy_file, index=False)

        # Create mock commission file with matching, duplicate, and unmatched rows
        comm_df = pd.DataFrame({
            "Policy No": ["POL100", "POL100", "POL200", "POL999"],
            "Product": ["Health", "Health", "Life", "Dental"],
            "Agent Name": ["Agent A", "Agent A", "Agent B", "Agent C"],
            "Sequence": [1, 1, 1, 1],
            "From Month": [1, 1, 1, 1],
            "Amount": ["$100.00", "$100.00", "$50.00", "$25.00"]
        })
        comm_file = os.path.join(self.temp_dir, "Commission_Data.csv")
        comm_df.to_csv(comm_file, index=False)

        out_dir = os.path.join(self.temp_dir, "output")

        valid_keys = extract_policy_keys_from_input(policy_file)
        self.assertEqual(valid_keys, {"POL100", "POL200", "POL300"})

        matched_dfs, unmatched_dfs = process_commission_files(
            comm_input_path=comm_file,
            valid_policy_keys=valid_keys,
            output_dir=out_dir,
            dedupe=True,
            separate_unmatched=True,
        )

        combined_matched = pd.concat(matched_dfs, ignore_index=True)
        combined_unmatched = pd.concat(unmatched_dfs, ignore_index=True)

        self.assertEqual(len(combined_matched), 2)  # POL100 (deduped to 1) + POL200 (1)
        self.assertEqual(len(combined_unmatched), 1)  # POL999
        self.assertEqual(set(combined_matched["Policy No"]), {"POL100", "POL200"})
        self.assertEqual(list(combined_unmatched["Policy No"]), ["POL999"])


if __name__ == "__main__":
    unittest.main()
