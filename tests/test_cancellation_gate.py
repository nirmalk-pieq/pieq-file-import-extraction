#!/usr/bin/env python3
"""Tests for Issue 6: Correct Policy Cancellation Gate — Strict Backend Parity.

Derived from PolicyCancellationGateTest.kt and PolicyAdvanceActiveArf.kt
in pieq-bg-service-icm-fileprocessing.
"""

from __future__ import annotations

import sys
import unittest
from decimal import Decimal
from pathlib import Path

# Add scripts/forecast to sys.path to import review_to_reconcile_file_forecast
FORECAST_DIR = Path(__file__).resolve().parent.parent / "scripts" / "forecast"
if str(FORECAST_DIR) not in sys.path:
    sys.path.insert(0, str(FORECAST_DIR))

import review_to_reconcile_file_forecast as forecast


class TestCancellationGate(unittest.TestCase):
    """Direct parity tests for Forecast.cancellation_gate and has_active_arf."""

    def setUp(self):
        self.fc = forecast.Forecast.__new__(forecast.Forecast)
        self.fc.exceptions = []

    def test_cancellation_gate_blocks_agency_active_arf_with_outstanding_balance(self):
        # Parity with PolicyCancellationGateTest:
        # `shouldFailEarlyPolicyCancellation for agency active ARF and chargeback`
        record = {
            "chargeback_amount": Decimal("-30.00"),
            "gross": Decimal("-30.00"),
            "has_policy_advance": True,
            "policy_advance_status": 1,
            "policy_advance_amount": Decimal("100.00"),
            "policy_advance_recovered_amount": Decimal("40.00"),
            "policy_advance_initiated_by": "agency",
            "is_policy_cancellation_validated": False,
        }
        self.assertTrue(self.fc.has_active_arf(record))
        self.assertTrue(self.fc.cancellation_gate(record))

    def test_cancellation_gate_does_not_block_when_advance_fully_recovered(self):
        # Parity with PolicyCancellationGateTest:
        # `shouldFailEarlyPolicyCancellation is false when advance is fully recovered`
        record = {
            "chargeback_amount": Decimal("-30.00"),
            "gross": Decimal("-30.00"),
            "has_policy_advance": True,
            "policy_advance_status": 1,
            "policy_advance_amount": Decimal("100.00"),
            "policy_advance_recovered_amount": Decimal("100.00"),  # Fully recovered!
            "policy_advance_initiated_by": "agency",
            "is_policy_cancellation_validated": False,
        }
        self.assertFalse(self.fc.has_active_arf(record))
        self.assertFalse(self.fc.cancellation_gate(record))

    def test_cancellation_gate_does_not_block_when_advance_over_recovered(self):
        record = {
            "chargeback_amount": Decimal("-30.00"),
            "gross": Decimal("-30.00"),
            "has_policy_advance": True,
            "policy_advance_status": 1,
            "policy_advance_amount": Decimal("100.00"),
            "policy_advance_recovered_amount": Decimal("120.00"),
            "policy_advance_initiated_by": "agency",
            "is_policy_cancellation_validated": False,
        }
        self.assertFalse(self.fc.has_active_arf(record))
        self.assertFalse(self.fc.cancellation_gate(record))

    def test_cancellation_gate_does_not_block_when_no_active_arf_status_not_one(self):
        # Status is not 1 (e.g. 0 or 2)
        record = {
            "chargeback_amount": Decimal("-30.00"),
            "gross": Decimal("-30.00"),
            "has_policy_advance": True,
            "policy_advance_status": 0,
            "policy_advance_amount": Decimal("100.00"),
            "policy_advance_recovered_amount": Decimal("0.00"),
            "policy_advance_initiated_by": "agency",
            "is_policy_cancellation_validated": False,
        }
        self.assertFalse(self.fc.has_active_arf(record))
        self.assertFalse(self.fc.cancellation_gate(record))

    def test_cancellation_gate_does_not_block_when_no_policy_advance(self):
        record = {
            "chargeback_amount": Decimal("-30.00"),
            "gross": Decimal("-30.00"),
            "has_policy_advance": False,
            "policy_advance_status": 1,
            "policy_advance_amount": Decimal("100.00"),
            "policy_advance_recovered_amount": Decimal("0.00"),
            "policy_advance_initiated_by": "agency",
            "is_policy_cancellation_validated": False,
        }
        self.assertFalse(self.fc.has_active_arf(record))
        self.assertFalse(self.fc.cancellation_gate(record))

    def test_cancellation_gate_does_not_block_for_carrier_initiated_advance(self):
        # Parity with PolicyCancellationGateTest:
        # `shouldFailEarlyPolicyCancellation is false for carrier initiated advance`
        record = {
            "chargeback_amount": Decimal("-30.00"),
            "gross": Decimal("-30.00"),
            "has_policy_advance": True,
            "policy_advance_status": 1,
            "policy_advance_amount": Decimal("100.00"),
            "policy_advance_recovered_amount": Decimal("0.00"),
            "policy_advance_initiated_by": "carrier",
            "is_policy_cancellation_validated": False,
        }
        self.assertTrue(self.fc.has_active_arf(record))
        self.assertFalse(self.fc.cancellation_gate(record))

    def test_cancellation_gate_does_not_block_when_policy_cancellation_validated(self):
        # Parity with PolicyCancellationGateTest:
        # `shouldFailEarlyPolicyCancellation is false when policy cancellation validated`
        record = {
            "chargeback_amount": Decimal("-30.00"),
            "gross": Decimal("-30.00"),
            "has_policy_advance": True,
            "policy_advance_status": 1,
            "policy_advance_amount": Decimal("100.00"),
            "policy_advance_recovered_amount": Decimal("0.00"),
            "policy_advance_initiated_by": "agency",
            "is_policy_cancellation_validated": True,
        }
        self.assertTrue(self.fc.has_active_arf(record))
        self.assertFalse(self.fc.cancellation_gate(record))

    def test_cancellation_gate_does_not_block_when_not_chargeback(self):
        # Parity with PolicyCancellationGateTest:
        # `isChargebackRecord is false for positive gross without chargeback`
        record = {
            "chargeback_amount": Decimal("0.00"),
            "gross": Decimal("30.00"),
            "has_policy_advance": True,
            "policy_advance_status": 1,
            "policy_advance_amount": Decimal("100.00"),
            "policy_advance_recovered_amount": Decimal("0.00"),
            "policy_advance_initiated_by": "agency",
            "is_policy_cancellation_validated": False,
        }
        self.assertTrue(self.fc.has_active_arf(record))
        self.assertFalse(self.fc.cancellation_gate(record))

    def test_reconciliation_fails_with_policy_cancellation_when_gate_triggers(self):
        record = {
            "uuid": "rec-cancel-1",
            "policy_number": "POL-1",
            "policy_exists": True,
            "policy_status_id": forecast.POLICY_ACTIVE,
            "agency_id": 1,
            "agent_npn": "12345",
            "carrier_agent_npn": "12345",
            "policy_uuid": "pol-1",
            "gross": Decimal("-50.00"),
            "chargeback_amount": Decimal("-50.00"),
            "has_policy_advance": True,
            "policy_advance_status": 1,
            "policy_advance_amount": Decimal("200.00"),
            "policy_advance_recovered_amount": Decimal("50.00"),
            "policy_advance_initiated_by": "agency",
            "is_policy_cancellation_validated": False,
        }
        self.fc.check_agent_mismatch = lambda r: False
        self.fc.maybe_reattach_commissions = lambda r: None

        self.fc.reconcile_record(record)

        self.assertEqual(len(self.fc.exceptions), 1)
        exc = self.fc.exceptions[0]
        self.assertEqual(exc["code"], "POLICY_CANCELLATION")
        self.assertEqual(exc["message"], "Policy cancellation required")


if __name__ == "__main__":
    unittest.main()