#!/usr/bin/env python3
"""Tests for forecast Policies_To_Update builder and writer."""

from __future__ import annotations

import csv
import sys
import tempfile
import unittest
from pathlib import Path

FORECAST_DIR = Path(__file__).resolve().parent.parent / "scripts" / "forecast"
if str(FORECAST_DIR) not in sys.path:
    sys.path.insert(0, str(FORECAST_DIR))

import review_to_reconcile_file_forecast as forecast

APL_DIR = Path(__file__).resolve().parent.parent / "scripts" / "apl_icm_cross_validation"
sys.path.insert(0, str(APL_DIR))
import validate_apl_icm as apl  # noqa: E402


class TestBuildPoliciesToUpdateRows(unittest.TestCase):
    def test_categories_agents_and_amounts(self):
        rows = forecast.build_policies_to_update_rows(
            [
                {
                    "ICM Policy Number": "POL-A",
                    "APL Agent Name": "Jane Doe",
                    "ICM Agent Name": "Jane Doe",
                    "APL Amount": "150.00",
                    "ICM Amount": "120.00",
                    "source_file_name": "file1.xlsx",
                    "Error": "Amount mismatch (APL: 150.0, ICM: 120.0)",
                },
                {
                    "ICM Policy Number": "POL-B",
                    "APL Agent Name": "Blair Agent",
                    "ICM Agent Name": "Blair A",
                    "APL Amount": "40.00",
                    "ICM Amount": "40.00",
                    "source_file_name": "file1.xlsx",
                    "Error": "Agent name does not match",
                },
                {
                    "ICM Policy Number": "POL-C",
                    "APL Agent Name": "Sam Lee",
                    "ICM Agent Name": "Sam L",
                    "APL Amount": "75.00",
                    "ICM Amount": "50.00",
                    "source_file_name": "file1.xlsx",
                    "Error": "Agent name does not match; Amount mismatch (APL: 75.0, ICM: 50.0)",
                },
            ],
            apl.clean_policy_number,
            apl.clean_amount,
        )
        by_policy = {row["Policy Number"]: row for row in rows}
        self.assertEqual(by_policy["POL-A"]["Mismatch Category"], "Amount_Only")
        self.assertEqual(by_policy["POL-A"]["Difference"], -30.0)
        self.assertIn("Amount mismatch", by_policy["POL-A"]["Issue"])
        self.assertEqual(by_policy["POL-B"]["Mismatch Category"], "Agent_Only")
        self.assertEqual(by_policy["POL-B"]["APL Agent Name"], "Blair Agent")
        self.assertEqual(by_policy["POL-B"]["ICM Agent Name"], "Blair A")
        self.assertIn("Agent name does not match", by_policy["POL-B"]["Issue"])
        self.assertEqual(by_policy["POL-C"]["Mismatch Category"], "Both")

    def test_issue_when_totals_match_but_line_amounts_differ(self):
        """A78254-style: aggregated APL/ICM totals match but line splits do not."""
        rows = forecast.build_policies_to_update_rows(
            [
                {
                    "ICM Policy Number": "A78254",
                    "APL Agent Name": "DANNA GREEN",
                    "ICM Agent Name": "Danna Green",
                    "APL Amount": "0.45",
                    "ICM Amount": "1.05",
                    "source_file_name": "2026-10-31_Elevanceaca.xlsx",
                    "Error": "Amount mismatch (APL DANNA GREEN: 0.45, ICM Danna Green: 1.05)",
                },
                {
                    "ICM Policy Number": "A78254",
                    "APL Agent Name": "DANNA GREEN",
                    "ICM Agent Name": "Danna Green",
                    "APL Amount": "12.6",
                    "ICM Amount": "9.45",
                    "source_file_name": "2026-10-31_Elevanceaca.xlsx",
                    "Error": "Amount mismatch (APL DANNA GREEN: 12.6, ICM Danna Green: 9.45)",
                },
                {
                    "ICM Policy Number": "A78254",
                    "APL Agent Name": "MAIN LINE BENEFITS",
                    "ICM Agent Name": "MLB parallel run",
                    "APL Amount": "1.95",
                    "ICM Amount": "1.35",
                    "source_file_name": "2026-10-31_Elevanceaca.xlsx",
                    "Error": "Amount mismatch (APL MAIN LINE BENEFITS: 1.95, ICM MLB parallel run: 1.35)",
                },
                {
                    "ICM Policy Number": "A78254",
                    "APL Agent Name": "MARSHA SANDEROCK",
                    "ICM Agent Name": "Marsha Sanderock",
                    "APL Amount": "0.0",
                    "ICM Amount": "3.15",
                    "source_file_name": "2026-10-31_Elevanceaca.xlsx",
                    "Error": "Amount mismatch (APL MARSHA SANDEROCK: 0.0, ICM Marsha Sanderock: 3.15)",
                },
            ],
            apl.clean_policy_number,
            apl.clean_amount,
        )
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["Difference"], 0.0)
        self.assertEqual(row["Mismatch Category"], "Amount_Only")
        self.assertIn("Line-level amount mismatch (policy totals match)", row["Issue"])
        self.assertIn("Amount mismatch (APL DANNA GREEN: 0.45, ICM Danna Green: 1.05)", row["Issue"])

    def test_excludes_sign_only_mismatch(self):
        rows = forecast.build_policies_to_update_rows(
            [
                {
                    "ICM Policy Number": "POL-S",
                    "APL Agent Name": "A",
                    "ICM Agent Name": "A",
                    "APL Amount": "-3.01",
                    "ICM Amount": "3.01",
                    "source_file_name": "file1.xlsx",
                    "Error": "Amount sign mismatch (APL: -3.01, ICM: 3.01)",
                }
            ],
            apl.clean_policy_number,
            apl.clean_amount,
        )
        self.assertEqual(rows, [])

    def test_consolidates_by_policy_and_file(self):
        rows = forecast.build_policies_to_update_rows(
            [
                {
                    "ICM Policy Number": "POL-A",
                    "APL Agent Name": "Alex",
                    "ICM Agent Name": "Alex",
                    "APL Amount": "100",
                    "ICM Amount": "80",
                    "source_file_name": "base.xlsx",
                    "Error": "Amount mismatch (APL: 100, ICM: 80)",
                },
                {
                    "ICM Policy Number": "POL-A",
                    "APL Agent Name": "Alex",
                    "ICM Agent Name": "Alex",
                    "APL Amount": "10",
                    "ICM Amount": "5",
                    "source_file_name": "base.xlsx",
                    "Error": "Amount mismatch (APL: 10, ICM: 5)",
                },
                {
                    "ICM Policy Number": "POL-A",
                    "APL Agent Name": "Pat",
                    "ICM Agent Name": "Patricia",
                    "APL Amount": "10",
                    "ICM Amount": "10",
                    "source_file_name": "admin.xlsx",
                    "Error": "Agent name does not match",
                },
            ],
            apl.clean_policy_number,
            apl.clean_amount,
        )
        self.assertEqual(len(rows), 2)
        by_file = {row["File Name"]: row for row in rows}
        self.assertEqual(by_file["base.xlsx"]["APL Amount"], 110.0)
        self.assertEqual(by_file["base.xlsx"]["ICM Amount"], 85.0)
        self.assertEqual(by_file["admin.xlsx"]["Mismatch Category"], "Agent_Only")


class TestWritePoliciesToUpdate(unittest.TestCase):
    def test_summary_block(self):
        detail = [
            {
                "Policy Number": "POL-A",
                "Mismatch Category": "Amount_Only",
                "APL Agent Name": "Jane",
                "ICM Agent Name": "Jane",
                "APL Amount": 150.0,
                "ICM Amount": 120.0,
                "Difference": -30.0,
                "Issue": "Amount mismatch | Amount mismatch (APL: 150.0, ICM: 120.0)",
                "File Name": "file1.xlsx",
            },
            {
                "Policy Number": "POL-B",
                "Mismatch Category": "Agent_Only",
                "APL Agent Name": "Blair Agent",
                "ICM Agent Name": "Blair A",
                "APL Amount": 40.0,
                "ICM Amount": 40.0,
                "Difference": 0.0,
                "Issue": "Agent name does not match",
                "File Name": "file1.xlsx",
            },
        ]
        with tempfile.TemporaryDirectory() as tmp:
            dest = Path(tmp)
            forecast.write_policies_to_update(detail, dest, icm_policy_count=200)
            csv_path = dest / "Policies_To_Update.csv"
            xlsx_path = dest / "Policies_To_Update.xlsx"
            self.assertFalse(csv_path.exists())
            self.assertTrue(xlsx_path.is_file())
            from openpyxl import load_workbook

            wb = load_workbook(xlsx_path)
            sheet = wb.active
            values = {
                sheet.cell(r, 1).value: sheet.cell(r, 2).value
                for r in range(1, sheet.max_row + 1)
                if sheet.cell(r, 1).value
            }
            self.assertEqual(values.get("Total Policies:"), 200)
            self.assertEqual(values.get("Update required:"), 2)
            self.assertEqual(values.get("Update required %:"), "1%")

    def test_count_icm_policies_from_detail(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "commission_detail_report.csv"
            with path.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.DictWriter(handle, fieldnames=["policy_number", "amount"])
                writer.writeheader()
                writer.writerow({"policy_number": "POL-1", "amount": "1"})
                writer.writerow({"policy_number": "POL-1", "amount": "2"})
                writer.writerow({"policy_number": "POL-2", "amount": "3"})
            policies = forecast.count_icm_policies_from_detail(path, apl.clean_policy_number)
            self.assertEqual(policies, {"POL-1", "POL-2"})


if __name__ == "__main__":
    unittest.main()
