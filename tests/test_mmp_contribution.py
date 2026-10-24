#!/usr/bin/env python3
"""Tests for Issue 7: Add MMP Contribution Processing — Strict Backend Parity.

Derived from MarketingContributionEvaluatorTest.kt and CommissionProcessingService.kt
in pieq-bg-service-icm-comm-processing.
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


class TestMmpContribution(unittest.TestCase):
    """Parity tests for MMP contribution processing in review_to_reconcile_file_forecast.py."""

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
        self.fc.today = date(2026, 10, 24)
        self.fc.agent_mmp_configs = {}

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

    def _set_agent_mmp_config(
        self,
        agent_uuid="agent-1",
        enabled=True,
        earning_type="Commission",
        cap_amount=Decimal("500.00"),
        percentage=None,
    ):
        self.fc.agent_mmp_configs[str(agent_uuid)] = {
            "marketing_contribution": enabled,
            "earning_type": earning_type,
            "marketing_contribution_amount": cap_amount,
            "marketing_contribution_percentage": percentage,
        }

    # 1. Agent is ineligible for MMP (marketing_contribution is False)
    def test_agent_ineligible_when_flag_false(self):
        self._set_agent_mmp_config(enabled=False)
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("500.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        mmp_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP]
        self.assertEqual(len(mmp_details), 0)

        comm_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_COMMISSION]
        self.assertEqual(len(comm_details), 1)
        self.assertEqual(comm_details[0]["amount"], Decimal("500.00"))

    # 2. Agent is eligible and has percentage-based contribution configured (e.g. 50%)
    def test_eligible_with_percentage_based_contribution(self):
        # 50% of 100.00 = 50.00 deduction, cap is 500.00
        self._set_agent_mmp_config(percentage=Decimal("50.00"), cap_amount=Decimal("500.00"))
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("100.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        comm_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_COMMISSION][0]
        self.assertEqual(comm_detail["amount"], Decimal("100.00"))

        mmp_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP][0]
        self.assertEqual(mmp_detail["amount"], Decimal("-50.00"))
        self.assertEqual(mmp_detail["status"], "ready_for_payment")

        # Ledger check: debit entry (credit=False)
        mmp_ledger = [l for l in self.fc.ledger_details if l["transaction_type"] == forecast.TYPE_MMP][0]
        self.assertEqual(mmp_ledger["amount"], Decimal("50.00"))
        self.assertFalse(mmp_ledger["credit"])

    # 3. Agent is eligible and has default 100% contribution (null percentage defaults to 100%)
    def test_eligible_with_default_percentage_100(self):
        # Null percentage defaults to 100%, commission 100.00 -> 100.00 deduction
        self._set_agent_mmp_config(percentage=None, cap_amount=Decimal("500.00"))
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("100.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        mmp_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP][0]
        self.assertEqual(mmp_detail["amount"], Decimal("-100.00"))

    # 4. Contribution is zero or eligible earnings are zero
    def test_zero_contribution_or_zero_earnings(self):
        self._set_agent_mmp_config()
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("0.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        mmp_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP]
        self.assertEqual(len(mmp_details), 0)

    # 5. Monthly cap has not been used (full cap available)
    def test_full_cap_available(self):
        # Cap 500.00, Commission 300.00, 100% -> Deducts 300.00
        self._set_agent_mmp_config(cap_amount=Decimal("500.00"), percentage=None)
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("300.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        mmp_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP][0]
        self.assertEqual(mmp_detail["amount"], Decimal("-300.00"))

    # 6. Existing MMP deductions partially consume the monthly cap
    def test_partial_cap_consumed(self):
        # Cap 500.00, previous deductions 400.00 -> remaining 100.00
        # Commission 200.00 (100% = 200.00) -> capped at remaining 100.00
        self._set_agent_mmp_config(cap_amount=Decimal("500.00"), percentage=None)
        # Seed an existing in-memory MMP deduction of 400.00
        self.fc.details.append({
            "uuid": "mmp-prev",
            "agency_id": 100,
            "entity_uuid": "agent-1",
            "commission_type": forecast.TYPE_MMP,
            "amount": Decimal("-400.00"),
        })

        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("200.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        new_mmp = [d for d in self.fc.details if d["uuid"] != "mmp-prev" and d["commission_type"] == forecast.TYPE_MMP]
        self.assertEqual(len(new_mmp), 1)
        self.assertEqual(new_mmp[0]["amount"], Decimal("-100.00"))

    # 7. Existing MMP deductions have exhausted the monthly cap (NoOffset)
    def test_exhausted_monthly_cap_no_offset(self):
        # Cap 500.00, previous deductions 500.00 -> remaining 0.00
        self._set_agent_mmp_config(cap_amount=Decimal("500.00"), percentage=None)
        self.fc.details.append({
            "uuid": "mmp-prev",
            "agency_id": 100,
            "entity_uuid": "agent-1",
            "commission_type": forecast.TYPE_MMP,
            "amount": Decimal("-500.00"),
        })

        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("200.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        new_mmp = [d for d in self.fc.details if d["uuid"] != "mmp-prev" and d["commission_type"] == forecast.TYPE_MMP]
        self.assertEqual(len(new_mmp), 0)

        comm_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_COMMISSION][0]
        self.assertEqual(comm_detail["amount"], Decimal("200.00"))

    # 8. Calculated contribution exceeds remaining cap (e.g. 50% is 100.00, remaining cap is 30.00)
    def test_contribution_exceeds_remaining_cap(self):
        self._set_agent_mmp_config(cap_amount=Decimal("30.00"), percentage=Decimal("50.00"))
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("200.00")  # 50% is 100.00, but cap is 30.00

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        mmp_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP][0]
        self.assertEqual(mmp_detail["amount"], Decimal("-30.00"))

    # 9. Advance recovery occurs before MMP calculation (MMP calculated on remainder)
    def test_advance_recovery_occurs_before_mmp(self):
        # Advance balance: 300.00. Gross split: 500.00.
        # Advance recovery: 300.00 (ADVANCE_EARNED).
        # Remaining commission: 200.00 (COMMISSION).
        # MMP configured: 50% of remaining commission = 100.00 (NOT 50% of 500.00 = 250.00!)
        self.fc.advance_ledgers.append({
            "uuid": "adv-1",
            "agency_id": 100,
            "policy_uuid": "pol-1",
            "agent_uuid": "agent-1",
            "entity_uuid": "agent-1",
            "amount": Decimal("300.00"),
            "transaction_id": "ADV-1",
        })
        self._set_agent_mmp_config(percentage=Decimal("50.00"), cap_amount=Decimal("500.00"))

        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("500.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        adv_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_ADVANCE_EARNED][0]
        self.assertEqual(adv_detail["amount"], Decimal("300.00"))

        comm_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_COMMISSION][0]
        self.assertEqual(comm_detail["amount"], Decimal("200.00"))

        mmp_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP][0]
        self.assertEqual(mmp_detail["amount"], Decimal("-100.00"))  # 50% of 200.00

    # 10. Advance recovery entries are excluded from MMP cap/deduction calculations
    def test_advance_recovery_excluded_from_mmp_cap_deductions(self):
        # An advance recovery of -300.00 in advance_ledgers must NOT count as an MMP deduction
        self.fc.advance_ledgers.append({
            "uuid": "adv-rec",
            "agency_id": 100,
            "policy_uuid": "pol-1",
            "agent_uuid": "agent-1",
            "entity_uuid": "agent-1",
            "amount": Decimal("-300.00"),
            "transaction_id": "ADV-REC",
        })
        self._set_agent_mmp_config(cap_amount=Decimal("500.00"), percentage=None)

        deductions = self.fc.sum_marketing_contribution_deductions(
            100, "agent-1", date(2026, 10, 1), date(2026, 10, 31)
        )
        self.assertEqual(deductions, Decimal("0.00"))

    # 11. Earning type matching is case-insensitive and filters correctly
    def test_earning_type_matching(self):
        # Configured for "Bonus" only -> COMMISSION does not trigger MMP
        self._set_agent_mmp_config(earning_type="Bonus")
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("500.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        mmp_details = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP]
        self.assertEqual(len(mmp_details), 0)

        # Configured for "Commission, Bonus" -> matches COMMISSION
        self._set_agent_mmp_config(earning_type="Commission, Bonus")
        self.fc.apply_commission(record, "event-2", match, gross, forecast.TYPE_COMMISSION)
        mmp_details2 = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP]
        self.assertEqual(len(mmp_details2), 1)

    # 12. Percentage out of range (0 or >100) returns NotApplicable
    def test_invalid_percentages_return_no_mmp(self):
        self._set_agent_mmp_config(percentage=Decimal("0.00"))
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("100.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)
        self.assertEqual(len([d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP]), 0)

        self._set_agent_mmp_config(percentage=Decimal("150.00"))
        self.fc.apply_commission(record, "event-2", match, gross, forecast.TYPE_COMMISSION)
        self.assertEqual(len([d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP]), 0)

    # 13. Detail and negative ledger entry agree on amount and period
    def test_detail_and_ledger_agreement(self):
        self._set_agent_mmp_config(percentage=Decimal("25.00"), cap_amount=Decimal("500.00"))
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("400.00")  # 25% of 400.00 = 100.00

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        mmp_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP][0]
        self.assertEqual(mmp_detail["amount"], Decimal("-100.00"))
        self.assertEqual(mmp_detail["status"], "ready_for_payment")

        mmp_ledger = [l for l in self.fc.ledger_details if l["transaction_type"] == forecast.TYPE_MMP][0]
        self.assertEqual(mmp_ledger["amount"], Decimal("100.00"))
        self.assertFalse(mmp_ledger["credit"])
        self.assertEqual(mmp_ledger["transaction_entity_uuid"], mmp_detail["uuid"])
        self.assertEqual(mmp_ledger["transaction_date"], self.fc.today)

    # 14. No outstanding advance recovery does not change existing MMP behavior
    def test_no_advance_mmp_behavior_unchanged(self):
        self._set_agent_mmp_config(percentage=Decimal("100.00"), cap_amount=Decimal("500.00"))
        record = self._sample_record()
        match = self._sample_single_split_match()
        gross = Decimal("250.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        # No advance ledgers
        self.assertEqual(len(self.fc.advance_ledgers), 0)
        # MMP deducted exactly 250.00
        mmp_detail = [d for d in self.fc.details if d["commission_type"] == forecast.TYPE_MMP][0]
        self.assertEqual(mmp_detail["amount"], Decimal("-250.00"))

    # 15. Regression: Multi-recipient split with mixed MMP configs and agency split
    def test_multi_recipient_mixed_mmp(self):
        # Agent 1 has MMP enabled (50%), Agent 2 has MMP disabled, Agency has no MMP
        self._set_agent_mmp_config(agent_uuid="agent-1", enabled=True, percentage=Decimal("50.00"), cap_amount=Decimal("500.00"))
        self._set_agent_mmp_config(agent_uuid="agent-2", enabled=False)

        record = self._sample_record()
        match = {
            "payout_method": forecast.PAYOUT_PERCENT,
            "fee_amount": Decimal("100.00"),
            "commission_split": json.dumps([
                {"user_uuid": "agent-1", "order": 1, "entity": "agent", "role": "Agent", "commission_value": 50},
                {"user_uuid": "agent-2", "order": 2, "entity": "agent", "role": "Manager", "commission_value": 30},
                {"user_uuid": "agency-1", "order": 3, "entity": "agency", "role": "Agency", "commission_value": 20},
            ]),
        }
        gross = Decimal("1000.00")

        self.fc.apply_commission(record, "event-1", match, gross, forecast.TYPE_COMMISSION)

        # Agent 1 receives 500.00 commission, 250.00 MMP deduction
        a1_mmp = [d for d in self.fc.details if d["entity_uuid"] == "agent-1" and d["commission_type"] == forecast.TYPE_MMP]
        self.assertEqual(len(a1_mmp), 1)
        self.assertEqual(a1_mmp[0]["amount"], Decimal("-250.00"))

        # Agent 2 receives 300.00 commission, 0 MMP deduction
        a2_mmp = [d for d in self.fc.details if d["entity_uuid"] == "agent-2" and d["commission_type"] == forecast.TYPE_MMP]
        self.assertEqual(len(a2_mmp), 0)

        # Agency receives 200.00 commission, 0 MMP deduction
        agency_mmp = [d for d in self.fc.details if d["entity_uuid"] == "agency-1" and d["commission_type"] == forecast.TYPE_MMP]
        self.assertEqual(len(agency_mmp), 0)


if __name__ == "__main__":
    unittest.main()
