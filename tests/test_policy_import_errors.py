import os
import unittest
import tempfile
import shutil
import pandas as pd
from update_policy_import_with_validation import (
    separate_errors,
    enrich_policy_data_with_validation as enrich_policy_data_with_errors,
    is_value_missing
)


class TestPolicyImportErrors(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_is_value_missing(self):
        self.assertTrue(is_value_missing(None))
        self.assertTrue(is_value_missing(""))
        self.assertTrue(is_value_missing("   "))
        self.assertTrue(is_value_missing("[UNMATCHED]"))
        self.assertTrue(is_value_missing("unmatched"))
        self.assertTrue(is_value_missing("NaN"))
        self.assertTrue(is_value_missing(r"\N"))
        self.assertTrue(is_value_missing(r"\\N"))
        self.assertFalse(is_value_missing("Carrier A"))
        self.assertFalse(is_value_missing("AG12345"))

    def test_separate_errors(self):
        data = {
            "Policy No": ["P1", "P2", "P3", "P4"],
            "Agent Name": ["Alice", "Bob", "Charlie", "David"],
            "Carrier": ["Carrier A", None, "Carrier C", "[UNMATCHED]"],
            "Agent ID": ["AG101", "AG102", None, "  "]
        }
        df = pd.DataFrame(data)

        clean_df, error_df = separate_errors(df)

        # Row P1: Clean (Carrier A, AG101)
        # Row P2: Carrier missing (Carrier is None, AG102) -> Error: "Carrier is missing"
        # Row P3: Agent ID missing (Carrier C, Agent ID is None) -> Error: "Agent ID not matched"
        # Row P4: Both missing (Carrier is [UNMATCHED], Agent ID is "  ") -> Error: "Carrier is missing, Agent ID not matched"

        self.assertEqual(len(clean_df), 1)
        self.assertEqual(clean_df.iloc[0]["Policy No"], "P1")

        self.assertEqual(len(error_df), 3)

        p2_err = error_df[error_df["Policy No"] == "P2"].iloc[0]["Error"]
        p3_err = error_df[error_df["Policy No"] == "P3"].iloc[0]["Error"]
        p4_err = error_df[error_df["Policy No"] == "P4"].iloc[0]["Error"]

        self.assertEqual(p2_err, "Carrier is missing")
        self.assertEqual(p3_err, "Agent ID not matched")
        self.assertEqual(p4_err, "Carrier is missing, Agent ID not matched")

    def test_enrich_policy_data_with_errors_e2e(self):
        # Base file
        # Base file with an unmatched policy POL005
        base_df = pd.DataFrame({
            "Policy No": ["POL001", "POL002", "POL003", "POL004", "POL005"],
            "Agent Name": ["John Doe", "Jane Smith", "Bob White", "Alice Brown", "Unknown Agent"],
        })
        base_path = os.path.join(self.test_dir, "base_policy.xlsx")
        base_df.to_excel(base_path, index=False)

        # Master file
        master_df = pd.DataFrame({
            "Policy Number": ["POL001", "POL002", "POL003", "POL004"],
            "Agent Name": ["John Doe", "Jane Smith", "Bob White", "Alice Brown"],
            "Carrier": ["Travelers", "Progressive", None, "Hartford"],
            "Agent ID": ["AG001", None, "AG003", None],
            "LOB": ["Auto", "Home", "Commercial", "Auto"],
            "Pay Mode": ["Annual", "Monthly", "Annual", "Monthly"],
            "Premium": [1000, 1500, 2000, 800]
        })
        master_path = os.path.join(self.test_dir, "master_policy.xlsx")
        master_df.to_excel(master_path, index=False)

        out_path = os.path.join(self.test_dir, "Policy_Enriched.xlsx")
        err_path = os.path.join(self.test_dir, "Policy_Errors.xlsx")

        enrich_policy_data_with_errors(
            base_file_path=base_path,
            master_file_path=master_path,
            output_file_path=out_path,
            error_file_path=err_path
        )

        self.assertTrue(os.path.exists(out_path))
        self.assertTrue(os.path.exists(err_path))

        clean_result = pd.read_excel(out_path)
        error_result = pd.read_excel(err_path)

        # POL001: Travelers, AG001 -> Clean
        self.assertEqual(len(clean_result), 1)
        self.assertEqual(clean_result.iloc[0]["Policy No"], "POL001")

        # POL002: Progressive, None -> "Agent ID not matched"
        # POL003: None, AG003 -> "Carrier is missing"
        # POL004: Hartford, None -> "Agent ID not matched"
        # POL005: Not in master file -> "Policy Not found, Carrier is missing, Agent ID not matched"
        self.assertEqual(len(error_result), 4)
        self.assertIn("Error", error_result.columns)
        
        pol005_err = error_result[error_result["Policy No"] == "POL005"].iloc[0]["Error"]
        self.assertIn("Policy Not found", pol005_err)

    def test_ref_file_unmatched_agent(self):
        # Base file with agent not present in ref_file
        base_df = pd.DataFrame({
            "Policy No": ["POL101", "POL102"],
            "Agent Name": ["Known Agent", "Unknown Agent"],
        })
        base_path = os.path.join(self.test_dir, "base_ref.xlsx")
        base_df.to_excel(base_path, index=False)

        master_df = pd.DataFrame({
            "Policy Number": ["POL101", "POL102"],
            "Agent Name": ["Known Agent", "Unknown Agent"],
            "Carrier": ["Travelers", "Travelers"],
            "Agent ID": ["AG101", "AG102"]
        })
        master_path = os.path.join(self.test_dir, "master_ref.xlsx")
        master_df.to_excel(master_path, index=False)

        ref_df = pd.DataFrame({
            "Agent Name": ["Known Agent"],
            "Agent ID": ["ROSTER_AG101"]
        })
        ref_path = os.path.join(self.test_dir, "roster.xlsx")
        ref_df.to_excel(ref_path, index=False)

        out_path = os.path.join(self.test_dir, "ref_out.xlsx")
        err_path = os.path.join(self.test_dir, "ref_err.xlsx")

        enrich_policy_data_with_errors(
            base_file_path=base_path,
            master_file_path=master_path,
            output_file_path=out_path,
            error_file_path=err_path,
            ref_file_path=ref_path
        )

        clean_res = pd.read_excel(out_path)
        err_res = pd.read_excel(err_path)

        self.assertEqual(len(clean_res), 1)
        self.assertEqual(clean_res.iloc[0]["Policy No"], "POL101")
        self.assertEqual(clean_res.iloc[0]["Agent ID"], "ROSTER_AG101")

        self.assertEqual(len(err_res), 1)
        self.assertEqual(err_res.iloc[0]["Policy No"], "POL102")
        self.assertEqual(err_res.iloc[0]["Error"], "Agent ID not matched")


if __name__ == "__main__":
    unittest.main()
