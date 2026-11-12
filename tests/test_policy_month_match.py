#!/usr/bin/env python3
"""Parity tests for 1-based policy months and MATCH lookback/lookahead (fileprocessing 2026.11.03.01)."""

from __future__ import annotations

import sys
import unittest
from datetime import date
from pathlib import Path

FORECAST_DIR = Path(__file__).resolve().parent.parent / "scripts" / "forecast"
if str(FORECAST_DIR) not in sys.path:
    sys.path.insert(0, str(FORECAST_DIR))

from review_to_reconcile_file_forecast import (
    MATCH_SQL,
    policy_month_index,
    slab_covers_policy_month,
    slab_in_match_window,
    sql_policy_month,
)


class TestPolicyMonthIndex(unittest.TestCase):
    """Parity with CommissionReattachment.policyRenewalMonthIndex at tag 2026.11.03.01."""

    def test_same_day_is_month_1(self):
        d = date(2025, 10, 15)
        self.assertEqual(policy_month_index(d, d), 1)

    def test_same_calendar_month_after_effective_is_month_1(self):
        self.assertEqual(
            policy_month_index(date(2025, 10, 20), date(2025, 10, 15)),
            1,
        )

    def test_next_month_is_month_2(self):
        self.assertEqual(
            policy_month_index(date(2025, 11, 1), date(2025, 10, 15)),
            2,
        )

    def test_boundary_month_13(self):
        self.assertEqual(
            policy_month_index(date(2026, 10, 31), date(2025, 10, 15)),
            13,
        )

    def test_boundary_month_14(self):
        self.assertEqual(
            policy_month_index(date(2026, 11, 1), date(2025, 10, 15)),
            14,
        )

    def test_paid_before_effective_is_none(self):
        self.assertIsNone(
            policy_month_index(date(2025, 10, 10), date(2025, 10, 15)),
        )

    def test_cross_year_december_to_january(self):
        self.assertEqual(
            policy_month_index(date(2026, 1, 15), date(2025, 12, 15)),
            2,
        )

    def test_sql_policy_month_matches_index(self):
        paid = date(2026, 10, 1)
        effective = date(2025, 10, 1)
        self.assertEqual(sql_policy_month(paid, effective), 13)
        self.assertEqual(policy_month_index(paid, effective), 13)


class TestSlabMatchWindow(unittest.TestCase):
    def test_exact_cover(self):
        regular = {"month_from": 1, "month_to": 12}
        self.assertTrue(slab_covers_policy_month(regular, 12))
        self.assertTrue(slab_in_match_window(regular, 12))

    def test_lookback_at_entry_into_renewal(self):
        regular = {"month_from": 1, "month_to": 12}
        renewal = {"month_from": 13, "month_to": None}
        # Policy month 13: exact renewal, lookback also includes regular (m-1 = 12)
        self.assertTrue(slab_in_match_window(renewal, 13))
        self.assertTrue(slab_in_match_window(regular, 13))
        self.assertFalse(slab_covers_policy_month(regular, 13))

    def test_lookahead_at_exit_from_regular(self):
        regular = {"month_from": 1, "month_to": 12}
        renewal = {"month_from": 13, "month_to": None}
        # Policy month 12: exact regular, lookahead also includes renewal (m+1 = 13)
        self.assertTrue(slab_in_match_window(regular, 12))
        self.assertTrue(slab_in_match_window(renewal, 12))
        self.assertFalse(slab_covers_policy_month(renewal, 12))

    def test_no_lookback_for_month_1(self):
        prior = {"month_from": 0, "month_to": 0}
        self.assertFalse(slab_in_match_window(prior, 1))


class TestMatchSqlShape(unittest.TestCase):
    def test_match_sql_has_1_based_and_window(self):
        self.assertIn("+ 1", MATCH_SQL)
        self.assertIn("calc.m - 1", MATCH_SQL)
        self.assertIn("calc.m + 1", MATCH_SQL)
        self.assertIn("paid_p >= params.effective_p", MATCH_SQL)
        self.assertEqual(MATCH_SQL.count("%s"), 4)


if __name__ == "__main__":
    unittest.main()
