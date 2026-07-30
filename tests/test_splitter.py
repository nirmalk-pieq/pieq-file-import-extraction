import os
import unittest
import tempfile
import shutil
import pandas as pd
from split_file import sanitize_filename, split_file
from extract_unique import extract_unique
from update_columns import update_columns
from process_effective_date import process_effective_date, calculate_month_range

class TestFileSplitter(unittest.TestCase):
    def setUp(self):
        # Create a temporary directory for test files
        self.test_dir = tempfile.mkdtemp()
        
        # Prepare mock data: 10 rows total
        # 5 rows with carrier a, 3 carrier b, 2 carrier c
        self.mock_data = {
            "Policy ID": [f"POL{i}" for i in range(1, 11)],
            "Carrier Name": [
                "Carrier A", "Carrier A", "Carrier A", "Carrier A", "Carrier A",
                "Carrier B", "Carrier B", "Carrier B",
                "Carrier C", "Carrier C"
            ],
            "Premium": [100 * i for i in range(1, 11)]
        }
        self.df = pd.DataFrame(self.mock_data)

    def tearDown(self):
        # Clean up temporary directory
        shutil.rmtree(self.test_dir)

    def test_sanitize_filename(self):
        self.assertEqual(sanitize_filename("Carrier A"), "Carrier A")
        self.assertEqual(sanitize_filename("Carrier/B?"), "Carrier_B_")
        self.assertEqual(sanitize_filename("  Carrier:C  "), "Carrier_C")
        self.assertEqual(sanitize_filename(""), "unassigned_group")

    def test_split_csv(self):
        # Save mock data as CSV
        input_csv = os.path.join(self.test_dir, "input.csv")
        self.df.to_csv(input_csv, index=False)

        # Output directory
        output_dir = os.path.join(self.test_dir, "output_csv")

        # Run splitting
        split_file(
            file_path=input_csv,
            column_name="Carrier Name",
            output_dir=output_dir,
            out_format="match"
        )

        # Verify output files
        expected_files = ["Carrier A.csv", "Carrier B.csv", "Carrier C.csv"]
        for filename in expected_files:
            file_path = os.path.join(output_dir, filename)
            self.assertTrue(os.path.exists(file_path), f"{filename} does not exist")
            
            # Read and verify content
            df_split = pd.read_csv(file_path)
            if "Carrier A" in filename:
                self.assertEqual(len(df_split), 5)
                self.assertTrue((df_split["Carrier Name"] == "Carrier A").all())
            elif "Carrier B" in filename:
                self.assertEqual(len(df_split), 3)
                self.assertTrue((df_split["Carrier Name"] == "Carrier B").all())
            elif "Carrier C" in filename:
                self.assertEqual(len(df_split), 2)
                self.assertTrue((df_split["Carrier Name"] == "Carrier C").all())

    def test_split_excel(self):
        # Save mock data as Excel
        input_xlsx = os.path.join(self.test_dir, "input.xlsx")
        self.df.to_excel(input_xlsx, index=False)

        # Output directory
        output_dir = os.path.join(self.test_dir, "output_excel")

        # Run splitting
        split_file(
            file_path=input_xlsx,
            column_name="Carrier Name",
            output_dir=output_dir,
            out_format="match"
        )

        # Verify output files
        expected_files = ["Carrier A.xlsx", "Carrier B.xlsx", "Carrier C.xlsx"]
        for filename in expected_files:
            file_path = os.path.join(output_dir, filename)
            self.assertTrue(os.path.exists(file_path), f"{filename} does not exist")
            
            # Read and verify content
            df_split = pd.read_excel(file_path)
            if "Carrier A" in filename:
                self.assertEqual(len(df_split), 5)
                self.assertTrue((df_split["Carrier Name"] == "Carrier A").all())
            elif "Carrier B" in filename:
                self.assertEqual(len(df_split), 3)
                self.assertTrue((df_split["Carrier Name"] == "Carrier B").all())
            elif "Carrier C" in filename:
                self.assertEqual(len(df_split), 2)
                self.assertTrue((df_split["Carrier Name"] == "Carrier C").all())

    def test_split_excel_to_csv(self):
        # Save mock data as Excel, but request CSV outputs
        input_xlsx = os.path.join(self.test_dir, "input.xlsx")
        self.df.to_excel(input_xlsx, index=False)

        output_dir = os.path.join(self.test_dir, "output_excel_to_csv")

        split_file(
            file_path=input_xlsx,
            column_name="Carrier Name",
            output_dir=output_dir,
            out_format="csv"
        )

        # Verify output files are CSVs
        expected_files = ["Carrier A.csv", "Carrier B.csv", "Carrier C.csv"]
        for filename in expected_files:
            file_path = os.path.join(output_dir, filename)
            self.assertTrue(os.path.exists(file_path))

    def test_missing_value_handling(self):
        # Add a row with NaN carrier
        df_with_nan = pd.concat([self.df, pd.DataFrame([{
            "Policy ID": "POL11",
            "Carrier Name": None,
            "Premium": 1100
        }])], ignore_index=True)

        input_csv = os.path.join(self.test_dir, "input_nan.csv")
        df_with_nan.to_csv(input_csv, index=False)

        output_dir = os.path.join(self.test_dir, "output_nan")

        split_file(
            file_path=input_csv,
            column_name="Carrier Name",
            output_dir=output_dir,
            out_format="match"
        )

        # Verify NaN values went to "unassigned_Carrier Name.csv"
        expected_nan_file = os.path.join(output_dir, "unassigned_Carrier Name.csv")
        self.assertTrue(os.path.exists(expected_nan_file))
        df_nan = pd.read_csv(expected_nan_file)
        self.assertEqual(len(df_nan), 1)
        self.assertEqual(df_nan.iloc[0]["Policy ID"], "POL11")

    def test_extract_unique_csv(self):
        agents_data = {
            "Agent ID": ["A1", "A2", "A3", "A4", "A5"],
            "Agent Name": ["Alice", "Bob", "Alice", None, "Bob"],
            "Email": ["alice@mail.com", "bob@mail.com", "alice2@mail.com", "unknown@mail.com", "bob2@mail.com"]
        }
        df_agents = pd.DataFrame(agents_data)
        input_csv = os.path.join(self.test_dir, "agents.csv")
        df_agents.to_csv(input_csv, index=False)

        output_csv = os.path.join(self.test_dir, "agents_unique.csv")
        
        extract_unique(
            file_path=input_csv,
            column_name="Agent Name",
            select_cols=["Agent Name", "Email"],
            output_path=output_csv
        )

        self.assertTrue(os.path.exists(output_csv))
        df_unique = pd.read_csv(output_csv)
        self.assertEqual(len(df_unique), 2)
        self.assertCountEqual(df_unique["Agent Name"].tolist(), ["Alice", "Bob"])
        self.assertCountEqual(df_unique.columns.tolist(), ["Agent Name", "Email"])

        unassigned_csv = os.path.join(self.test_dir, "agents_unique_unassigned.csv")
        self.assertTrue(os.path.exists(unassigned_csv))
        df_unassigned = pd.read_csv(unassigned_csv)
        self.assertEqual(len(df_unassigned), 1)
        self.assertEqual(df_unassigned.iloc[0]["Email"], "unknown@mail.com")

    def test_update_columns_single_file(self):
        # Reference source file
        df_source = pd.DataFrame({
            "Agent Name": ["Alice", "Bob", "Charlie"],
            "Agent ID": ["AG101", "AG102", "AG103"]
        })
        source_xlsx = os.path.join(self.test_dir, "Agent_data.xlsx")
        df_source.to_excel(source_xlsx, index=False)

        # Target file missing Agent ID
        df_target = pd.DataFrame({
            "Agent Name": ["Alice", "Bob", "Alice"],
            "Carrier": ["Carrier A", "Carrier B", "Carrier C"]
        })
        target_xlsx = os.path.join(self.test_dir, "policies.xlsx")
        df_target.to_excel(target_xlsx, index=False)

        output_dir = os.path.join(self.test_dir, "policies_updated_dir")

        update_columns(
            source_path=source_xlsx,
            key_column="Agent Name",
            update_cols=["Agent ID"],
            target_path=target_xlsx,
            output_dir=output_dir
        )

        expected_out = os.path.join(output_dir, "policies_updated.xlsx")
        self.assertTrue(os.path.exists(expected_out))
        df_res = pd.read_excel(expected_out)
        self.assertIn("Agent ID", df_res.columns)
        self.assertEqual(df_res["Agent ID"].tolist(), ["AG101", "AG102", "AG101"])

    def test_update_columns_folder(self):
        # Reference source file
        df_source = pd.DataFrame({
            "Agent Name": ["Alice", "Bob"],
            "Agent ID": ["AG101", "AG102"]
        })
        source_csv = os.path.join(self.test_dir, "Agent_ref.csv")
        df_source.to_csv(source_csv, index=False)

        # Target directory with split files
        target_dir = os.path.join(self.test_dir, "split_folder")
        os.makedirs(target_dir, exist_ok=True)

        df_file1 = pd.DataFrame({"Agent Name": ["Alice"], "Policy": ["P1"]})
        df_file2 = pd.DataFrame({"Agent Name": ["Bob"], "Policy": ["P2"]})

        df_file1.to_csv(os.path.join(target_dir, "file1.csv"), index=False)
        df_file2.to_excel(os.path.join(target_dir, "file2.xlsx"), index=False)

        output_dir = os.path.join(self.test_dir, "split_folder_updated")

        update_columns(
            source_path=source_csv,
            key_column="Agent Name",
            update_cols=["Agent ID"],
            target_path=target_dir,
            output_dir=output_dir
        )

        res1_path = os.path.join(output_dir, "file1.csv")
        res2_path = os.path.join(output_dir, "file2.xlsx")

        self.assertTrue(os.path.exists(res1_path))
        self.assertTrue(os.path.exists(res2_path))

        res1 = pd.read_csv(res1_path)
        res2 = pd.read_excel(res2_path)

        self.assertEqual(res1.iloc[0]["Agent ID"], "AG101")
        self.assertEqual(res2.iloc[0]["Agent ID"], "AG102")

    def test_calculate_month_range_logic(self):
        from datetime import datetime
        ref_date = datetime(2026, 7, 24)

        # 4 months ago -> 1..12
        fm, tm = calculate_month_range("2026-03-01", ref_date)
        self.assertEqual(fm, 1)
        self.assertEqual(tm, 12)

        # 12 months ago -> 1..12
        fm, tm = calculate_month_range("2025-07-01", ref_date)
        self.assertEqual(fm, 1)
        self.assertEqual(tm, 12)

        # 13 months ago -> 13 & blank
        fm, tm = calculate_month_range("2025-06-01", ref_date)
        self.assertEqual(fm, 13)
        self.assertEqual(tm, "")

        # Missing date -> blank & blank
        fm, tm = calculate_month_range(None, ref_date)
        self.assertEqual(fm, "")
        self.assertEqual(tm, "")

    def test_process_effective_date_file(self):
        df = pd.DataFrame({
            "Policy ID": ["POL1", "POL2", "POL3"],
            "Effective Date": ["2026-03-01", "2025-05-01", None]
        })
        input_xlsx = os.path.join(self.test_dir, "eff_test.xlsx")
        df.to_excel(input_xlsx, index=False)

        output_dir = os.path.join(self.test_dir, "eff_processed_dir")

        process_effective_date(
            input_path=input_xlsx,
            date_column="Effective Date",
            current_date_str="2026-07-24",
            output_dir=output_dir
        )

        out_path = os.path.join(output_dir, "eff_test_processed.xlsx")
        self.assertTrue(os.path.exists(out_path))

        df_res = pd.read_excel(out_path)
        self.assertIn("From Month", df_res.columns)
        self.assertIn("To Month", df_res.columns)

        # Row 0: 2026-03-01 -> From Month 1, To Month 12
        self.assertEqual(df_res.iloc[0]["From Month"], 1)
        self.assertEqual(df_res.iloc[0]["To Month"], 12)

        # Row 1: 2025-05-01 -> From Month 13, To Month blank/NaN
        self.assertEqual(df_res.iloc[1]["From Month"], 13)
        self.assertTrue(pd.isna(df_res.iloc[1]["To Month"]) or df_res.iloc[1]["To Month"] == "")

        # Row 2: None -> From Month blank/NaN, To Month blank/NaN
        self.assertTrue(pd.isna(df_res.iloc[2]["From Month"]) or df_res.iloc[2]["From Month"] == "")
        self.assertTrue(pd.isna(df_res.iloc[2]["To Month"]) or df_res.iloc[2]["To Month"] == "")

if __name__ == "__main__":
    unittest.main()
