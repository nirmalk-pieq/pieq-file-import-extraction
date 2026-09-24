#!/usr/bin/env python3
import os
import shutil
import tempfile
import unittest
import pandas as pd
from normalize_policy_products import process_product_normalization


class TestNormalizePolicyProducts(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_product_normalization_fallback(self):
        # Create mock policy dataframe
        policy_df = pd.DataFrame({
            "Policy No": ["POL001", "POL002", "POL003", "POL004"],
            "Product": ["Accident 2024", "10 Pay Solution", "Empty Target Prod", "Unknown Prod"],
            "Agent Name": ["Agent A", "Agent B", "Agent C", "Agent D"]
        })
        policy_file = os.path.join(self.temp_dir, "policy.xlsx")
        policy_df.to_excel(policy_file, index=False)

        # Create mock normalization dataframe
        norm_df = pd.DataFrame({
            "Product ": ["Accident 2024", "10 Pay Solution", "Empty Target Prod"],
            "New Product Name ": ["Accident Insurance", "10 Pay Plan", ""]  # Empty string for 3rd
        })
        norm_file = os.path.join(self.temp_dir, "norm.xlsx")
        norm_df.to_excel(norm_file, index=False)

        clean_path = process_product_normalization(
            policy_path=policy_file,
            norm_path=norm_file,
            output_dir=self.temp_dir
        )

        # Verify output file contains ALL 4 rows, using original Product when New Product Name is empty/missing
        df_clean = pd.read_excel(clean_path)
        self.assertEqual(len(df_clean), 4)
        self.assertEqual(
            df_clean["Product"].tolist(),
            ["Accident Insurance", "10 Pay Plan", "Empty Target Prod", "Unknown Prod"]
        )


if __name__ == "__main__":
    unittest.main()
