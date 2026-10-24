#!/usr/bin/env python3
"""Tests for Issue 1: Add Advance Recovery to Commission Processing — Strict Backend Parity.

Derived from AdvanceRecoveryProcessorTest.kt, MultiRecipientAdvanceRecoveryTest.kt,
and CommissionProcessingService.kt in pieq-bg-service-icm-comm-processing.
"""

from __future__ import annotations

import json
import sys
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

# Add scripts/forecast to sys.path to import review_to_reconcile_file_forecast
FORECAST_DIR = Path(__file__).resolve().parent.parent / "scripts" / "forecast"
if str(FORECAST_DIR) not in sys.path:
    sys.path.insert(0, str(FORECAST_DIR))

import review_to_reconcile_file_forecast as forecast


class TestAdvanceRecovery(unittest.TestCase):
    """Parity tests for advance recovery in review_to_reconcile_file_forecast.py."""

    def setUp(self):
        self.fc = forecast.Forecast.__new__(forecast.Forecast)
        self.fc.conn = None
        self.fc.advance_ledgers = []
        self.fc.details = []
        self.fc.ledgers = {}
        self.fc.ledger_details = []
        self.fc.commissions = []
        self.fc.exceptions = []
        self.fc.tx = 0
        self.fc.today = date.today()

    def _sample_record(self, agency_id=100, policy_uuid="pol-1", statement_uuid="rec-1", **kwargs):
        record = {
            "uuid": statement_uuid,
            "agency_id": agency_id,
            "forecast_policy_uuid": policy_uuid,
            "policy_uuid": policy_uuid,
            "policy_number": "POL-100",
            "policy_advance_uuid": "pa-1",
            "premium_amount": Decimal("1000.00"),
            "policy_status_id": 0,
        }
        record.update(kwargs)
        return record

    def _sample_single_split_match(self, agent_uuid="agent-1", percent=100):
        return {
            "payout_method": forecast.PAYOUT_PERCENT,
            "fee_amount": Decimal("100.00"),
            "commission_split": json.dumps([
                {
                    "user_uuid": agent_uuid,
                    "order": 1,
                    "entity": "agent",
                    "role": "Agent",
                    "commission_value": percent,
                }
            ]),
        }

    def _seed_advance(self, agency_id=100, policy_uuid="pol-1", agent_uuid="agent-1", amount=Decimal("300.00")):
        self.fc.advance_ledgers.append(
            {
                "uuid": forecast.new_id(),
                "agency_id": agency_id,
                "policy_uuid": policy_uuid,
                "agent_uuid": agent_uuid,
                "entity_uuid": agent_uuid,
                "amount": amount,
                "transaction_id": self.fc.next_tx("ADV"),
            }
        )

    # 1. Partial recovery: Outstanding advance is less than the eligible split.
    def test_partial_recovery(self):
        # Outstanding: 300.00, Eligible Split: 500.00
        # Expected: Recover 300.00 (ADVANCE_EARNED), Remaining 200.00 (COMMISSION)
        self._seed_advance(amount=Decimal("300.00"))
        record = self._sample_record()
        match = self._sample_single_split_match(percent=100)
        gross = Decimal("500.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        # Details: 1 ADVANCE_EARNED for 300.00, 1 COMMISSION for 200.00
        adv_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_ADVANCE_EARNED]
        comm_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_COMMISSION]

        self.assertEqual(len(adv_details), 1)
        self.assertEqual(adv_details[0]["amount"], Decimal("300.00"))
        self.assertEqual(adv_details[0]["status"], "paid")

        self.assertEqual(len(comm_details), 1)
        self.assertEqual(comm_details[0]["amount"], Decimal("200.00"))
        self.assertEqual(comm_details[0]["status"], "ready_for_payment")

        # Advance ledgers: initial 300.00 + recovery -300.00 -> balance 0.00
        rec_ledgers = [l for l in self.fc.advance_ledgers if l["amount"] < 0]
        self.assertEqual(len(rec_ledgers), 1)
        self.assertEqual(rec_ledgers[0]["amount"], Decimal("-300.00"))
        self.assertEqual(rec_ledgers[0]["agent_uuid"], "agent-1")

        outstanding_after = self.fc.get_agent_advance_outstanding(100, "pol-1", "agent-1")
        self.assertEqual(outstanding_after, Decimal("0.00"))

    # 2. Full split recovery: Outstanding advance equals the eligible split.
    def test_full_split_recovery(self):
        # Outstanding: 500.00, Eligible Split: 500.00
        # Expected: Recover 500.00 (ADVANCE_EARNED), 0 commission payable (no 0 detail)
        self._seed_advance(amount=Decimal("500.00"))
        record = self._sample_record()
        match = self._sample_single_split_match(percent=100)
        gross = Decimal("500.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        adv_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_ADVANCE_EARNED]
        comm_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_COMMISSION]

        self.assertEqual(len(adv_details), 1)
        self.assertEqual(adv_details[0]["amount"], Decimal("500.00"))
        self.assertEqual(adv_details[0]["status"], "paid")

        # No residual 0.00 COMMISSION detail
        self.assertEqual(len(comm_details), 0)

        # Advance ledger recovery entry
        rec_ledgers = [l for l in self.fc.advance_ledgers if l["amount"] < 0]
        self.assertEqual(len(rec_ledgers), 1)
        self.assertEqual(rec_ledgers[0]["amount"], Decimal("-500.00"))

        outstanding_after = self.fc.get_agent_advance_outstanding(100, "pol-1", "agent-1")
        self.assertEqual(outstanding_after, Decimal("0.00"))

    # 3. Advance exceeds split: Recover only the eligible split; preserve remaining advance.
    def test_advance_exceeds_split(self):
        # Outstanding: 600.00, Eligible Split: 500.00
        # Expected: Recover 500.00, Remaining advance balance: 100.00
        self._seed_advance(amount=Decimal("600.00"))
        record = self._sample_record()
        match = self._sample_single_split_match(percent=100)
        gross = Decimal("500.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        adv_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_ADVANCE_EARNED]
        comm_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_COMMISSION]

        self.assertEqual(len(adv_details), 1)
        self.assertEqual(adv_details[0]["amount"], Decimal("500.00"))
        self.assertEqual(len(comm_details), 0)

        # Remaining advance balance is 100.00
        outstanding_after = self.fc.get_agent_advance_outstanding(100, "pol-1", "agent-1")
        self.assertEqual(outstanding_after, Decimal("100.00"))

    # 4. Zero outstanding balance: Preserve existing commission allocation without recovery entry.
    def test_zero_outstanding_balance(self):
        # Outstanding: 0.00, Eligible Split: 500.00
        # Expected: Entire 500.00 is COMMISSION, no ADVANCE_EARNED detail or advance ledger entry
        record = self._sample_record()
        match = self._sample_single_split_match(percent=100)
        gross = Decimal("500.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        adv_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_ADVANCE_EARNED]
        comm_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_COMMISSION]

        self.assertEqual(len(adv_details), 0)
        self.assertEqual(len(comm_details), 1)
        self.assertEqual(comm_details[0]["amount"], Decimal("500.00"))
        self.assertEqual(comm_details[0]["status"], "ready_for_payment")

        # Zero advance ledger entries
        self.assertEqual(len(self.fc.advance_ledgers), 0)

    # 5. Zero eligible split: No invalid recovery created.
    def test_zero_eligible_split(self):
        # Outstanding: 300.00, Eligible Split: 0.00
        self._seed_advance(amount=Decimal("300.00"))
        record = self._sample_record()
        match = self._sample_single_split_match(percent=100)
        gross = Decimal("0.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        adv_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_ADVANCE_EARNED]
        self.assertEqual(len(adv_details), 0)

        # Advance remains unreduced
        outstanding_after = self.fc.get_agent_advance_outstanding(100, "pol-1", "agent-1")
        self.assertEqual(outstanding_after, Decimal("300.00"))

    # 6. Ledger/detail consistency: Verify recovery amount, classification, signs, and credits.
    def test_ledger_detail_consistency(self):
        self._seed_advance(amount=Decimal("400.00"))
        record = self._sample_record(policy_advance_recovered_amount=Decimal("0.00"))
        match = self._sample_single_split_match(percent=100)
        gross = Decimal("250.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        # Check detail: positive amount, status paid, commission_type ADVANCE_EARNED
        adv_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_ADVANCE_EARNED][0]
        self.assertEqual(adv_detail["amount"], Decimal("250.00"))
        self.assertEqual(adv_detail["status"], "paid")
        self.assertEqual(adv_detail["statement_uuid"], "rec-1")

        # Check advance ledger: negative amount, agent_uuid
        adv_ledger = [l for l in self.fc.advance_ledgers if l["amount"] < 0][0]
        self.assertEqual(adv_ledger["amount"], Decimal("-250.00"))
        self.assertEqual(adv_ledger["agent_uuid"], "agent-1")
        self.assertEqual(adv_ledger["policy_uuid"], "pol-1")
        self.assertEqual(adv_ledger["commission_statement_record_uuid"], "rec-1")

        # Check account ledger detail: transaction_type ADVANCE_EARNED, credit True
        acct_ledger = [l for l in self.fc.ledger_details if l["transaction_type"] == forecast.TYPE_ADVANCE_EARNED][0]
        self.assertEqual(acct_ledger["amount"], Decimal("250.00"))
        self.assertTrue(acct_ledger["credit"])
        self.assertEqual(acct_ledger["transaction_entity_uuid"], adv_detail["uuid"])

        # Check record recovered amount update
        self.assertEqual(record["policy_advance_recovered_amount"], Decimal("250.00"))

    # 7. Regression: Ordinary commission persistence and multi-tier allocation unchanged when no advance.
    def test_regression_multi_recipient_without_advance(self):
        record = self._sample_record()
        match = {
            "payout_method": forecast.PAYOUT_PERCENT,
            "fee_amount": Decimal("100.00"),
            "commission_split": json.dumps([
                {"user_uuid": "agent-1", "order": 1, "entity": "agent", "role": "Agent", "commission_value": 70},
                {"user_uuid": "agent-2", "order": 2, "entity": "agent", "role": "Manager", "commission_value": 30},
            ]),
        }
        gross = Decimal("1000.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        self.assertEqual(len(self.fc.advance_ledgers), 0)
        self.assertEqual(len(self.fc.details), 2)
        d1 = [d for d in self.fc.details if d["entity_uuid"] == "agent-1"][0]
        d2 = [d for d in self.fc.details if d["entity_uuid"] == "agent-2"][0]
        self.assertEqual(d1["amount"], Decimal("700.00"))
        self.assertEqual(d1["commission_type"], forecast.TYPE_COMMISSION)
        self.assertEqual(d2["amount"], Decimal("300.00"))
        self.assertEqual(d2["commission_type"], forecast.TYPE_COMMISSION)

    # 8. Multi-recipient advance recovery (Parity with MultiRecipientAdvanceRecoveryTest.kt).
    def test_multi_recipient_advance_recovery_agent_recovers_manager_does_not(self):
        # Agent 1 has 200.00 advance. Agent 2 has NO advance.
        # Split is 600.00 for Agent 1 (60%), 400.00 for Agent 2 (40%) on 1000.00 gross.
        self._seed_advance(agent_uuid="agent-1", amount=Decimal("200.00"))
        record = self._sample_record()
        match = {
            "payout_method": forecast.PAYOUT_PERCENT,
            "fee_amount": Decimal("100.00"),
            "commission_split": json.dumps([
                {"user_uuid": "agent-1", "order": 1, "entity": "agent", "role": "Agent", "commission_value": 60},
                {"user_uuid": "agent-2", "order": 2, "entity": "agent", "role": "Manager", "commission_value": 40},
            ]),
        }
        gross = Decimal("1000.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        # Agent 1: 200.00 recovered (ADVANCE_EARNED), 400.00 COMMISSION
        a1_adv = [d for d in self.fc.details if d["entity_uuid"] == "agent-1" and d["commission_type"] == forecast.TYPE_ADVANCE_EARNED]
        a1_comm = [d for d in self.fc.details if d["entity_uuid"] == "agent-1" and d["commission_type"] == forecast.TYPE_COMMISSION]
        self.assertEqual(len(a1_adv), 1)
        self.assertEqual(a1_adv[0]["amount"], Decimal("200.00"))
        self.assertEqual(len(a1_comm), 1)
        self.assertEqual(a1_comm[0]["amount"], Decimal("400.00"))

        # Agent 2: No recovery, full 400.00 COMMISSION
        a2_adv = [d for d in self.fc.details if d["entity_uuid"] == "agent-2" and d["commission_type"] == forecast.TYPE_ADVANCE_EARNED]
        a2_comm = [d for d in self.fc.details if d["entity_uuid"] == "agent-2" and d["commission_type"] == forecast.TYPE_COMMISSION]
        self.assertEqual(len(a2_adv), 0)
        self.assertEqual(len(a2_comm), 1)
        self.assertEqual(a2_comm[0]["amount"], Decimal("400.00"))

        # Agent 1 advance is now fully recovered (0.00)
        self.assertEqual(self.fc.get_agent_advance_outstanding(100, "pol-1", "agent-1"), Decimal("0.00"))

    # 9. Duplicate recovery prevention: has_recovery_ledger.
    def test_duplicate_recovery_prevention(self):
        self._seed_advance(amount=Decimal("300.00"))
        record = self._sample_record()

        # Check initial state
        self.assertFalse(self.fc.has_recovery_ledger(100, "pol-1", "agent-1", "rec-1"))

        # First persist recovery entry
        line = {"user_uuid": "agent-1", "entity": "agent", "role": "Agent", "amount": Decimal("200.00")}
        self.fc.persist_recovery_entry(record, "comm-1", line, Decimal("200.00"))

        # Now has_recovery_ledger should be True
        self.assertTrue(self.fc.has_recovery_ledger(100, "pol-1", "agent-1", "rec-1"))

        # Calling persist_recovery_entry again for the same statement record should not duplicate ledger
        self.fc.persist_recovery_entry(record, "comm-1", line, Decimal("200.00"))
        adv_ledgers_rec = [l for l in self.fc.advance_ledgers if l["amount"] < 0]
        self.assertEqual(len(adv_ledgers_rec), 1)


if __name__ == "__main__":
    unittest.main()
