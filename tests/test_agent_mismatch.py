#!/usr/bin/env python3
"""Tests for Issue 4: Agent NPN and Name Conflict Validation — Strict Backend Parity.

Derived from ValidateAgentTest.kt, ValidateAgentAndPolicyTest.kt, and PolicyCreationDaoTest.kt
in pieq-bg-service-icm-fileprocessing.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

# Add scripts/forecast to sys.path to import review_to_reconcile_file_forecast
FORECAST_DIR = Path(__file__).resolve().parent.parent / "scripts" / "forecast"
if str(FORECAST_DIR) not in sys.path:
    sys.path.insert(0, str(FORECAST_DIR))

import review_to_reconcile_file_forecast as forecast
from review_to_reconcile_file_forecast import (
    AGENT_MISMATCH_CODE,
    AGENT_MISMATCH_ID_NAME_MESSAGE,
    AGENT_NAME_RESOLVE_SQL,
    AGENT_RESOLVE_SQL,
    POLICY_ACTIVE,
    AgentNames,
    agent_matches,
)


class TestAgentMismatchResolutionHelpers(unittest.TestCase):
    """Tests resolve_agent_by_npn and resolve_agent_by_name against PolicyCreationDao.kt."""

    def setUp(self):
        self.fc = forecast.Forecast.__new__(forecast.Forecast)
        self.fc.exceptions = []

    def test_resolve_agent_by_npn_unique(self):
        # 1 distinct UUID -> Resolved
        self.fc.select = lambda sql, params: [
            {"agent_uuid": "agent-uuid-1", "resolution_method": "carrier_appointment", "priority": 2}
        ]
        uuid, status, rows = self.fc.resolve_agent_by_npn(1, "NPN123", "carrier-1")
        self.assertEqual(status, "Resolved")
        self.assertEqual(uuid, "agent-uuid-1")
        self.assertEqual(len(rows), 1)

    def test_resolve_agent_by_npn_multiple_branches_same_uuid(self):
        # Multiple rows returning the same UUID across branches -> Resolved (highest priority first)
        self.fc.select = lambda sql, params: [
            {"agent_uuid": "agent-uuid-1", "resolution_method": "npn", "priority": 1},
            {"agent_uuid": "agent-uuid-1", "resolution_method": "carrier_appointment", "priority": 2},
        ]
        uuid, status, rows = self.fc.resolve_agent_by_npn(1, "NPN123", "carrier-1")
        self.assertEqual(status, "Resolved")
        self.assertEqual(uuid, "agent-uuid-1")

    def test_resolve_agent_by_npn_conflicting_matches(self):
        # >1 distinct UUIDs -> ConflictingMatches
        self.fc.select = lambda sql, params: [
            {"agent_uuid": "agent-uuid-1", "resolution_method": "npn", "priority": 1},
            {"agent_uuid": "agent-uuid-2", "resolution_method": "user_code", "priority": 3},
        ]
        uuid, status, rows = self.fc.resolve_agent_by_npn(1, "NPN123")
        self.assertEqual(status, "ConflictingMatches")
        self.assertIsNone(uuid)
        self.assertEqual(len(rows), 2)

    def test_resolve_agent_by_npn_not_found(self):
        self.fc.select = lambda sql, params: []
        uuid, status, rows = self.fc.resolve_agent_by_npn(1, "UNKNOWN_NPN")
        self.assertEqual(status, "NotFound")
        self.assertIsNone(uuid)
        self.assertEqual(rows, [])

    def test_resolve_agent_by_npn_blank_or_none(self):
        self.fc.select = lambda sql, params: [
            {"agent_uuid": "uuid", "resolution_method": "npn", "priority": 1}
        ]
        for val in [None, "", "   "]:
            uuid, status, rows = self.fc.resolve_agent_by_npn(1, val)
            self.assertEqual(status, "NotFound")
            self.assertIsNone(uuid)

    def test_resolve_agent_by_name_unique(self):
        self.fc.select = lambda sql, params: [
            {"agent_uuid": "agent-uuid-name", "resolution_method": "full_name"}
        ]
        uuid, status, rows = self.fc.resolve_agent_by_name(1, "Blair Agent")
        self.assertEqual(status, "Resolved")
        self.assertEqual(uuid, "agent-uuid-name")

    def test_resolve_agent_by_name_ambiguous(self):
        # Multiple distinct contacts -> AmbiguousFullName
        self.fc.select = lambda sql, params: [
            {"agent_uuid": "agent-1", "resolution_method": "full_name"},
            {"agent_uuid": "agent-2", "resolution_method": "full_name"},
        ]
        uuid, status, rows = self.fc.resolve_agent_by_name(1, "John Smith")
        self.assertEqual(status, "AmbiguousFullName")
        self.assertIsNone(uuid)

    def test_resolve_agent_by_name_not_found(self):
        self.fc.select = lambda sql, params: []
        uuid, status, rows = self.fc.resolve_agent_by_name(1, "Ghost Agent")
        self.assertEqual(status, "NotFound")
        self.assertIsNone(uuid)

    def test_resolve_agent_by_name_blank_or_none(self):
        for val in [None, "", "   "]:
            uuid, status, rows = self.fc.resolve_agent_by_name(1, val)
            self.assertEqual(status, "NotFound")
            self.assertIsNone(uuid)


class TestCheckAgentMismatch(unittest.TestCase):
    """Tests Forecast.check_agent_mismatch behavior."""

    def setUp(self):
        self.fc = forecast.Forecast.__new__(forecast.Forecast)

    def test_mismatch_detected_when_both_resolve_to_different_uuids(self):
        self.fc.resolve_agent_by_npn = lambda aid, npn, cid=None: ("agent-A-uuid", "Resolved", [])
        self.fc.resolve_agent_by_name = lambda aid, name: ("agent-B-uuid", "Resolved", [])

        record = {"agency_id": 1, "agent_npn": "12345", "agent_name": "Agent B"}
        self.assertTrue(self.fc.check_agent_mismatch(record))

    def test_no_mismatch_when_both_resolve_to_same_uuid(self):
        self.fc.resolve_agent_by_npn = lambda aid, npn, cid=None: ("agent-A-uuid", "Resolved", [])
        self.fc.resolve_agent_by_name = lambda aid, name: ("agent-A-uuid", "Resolved", [])

        record = {"agency_id": 1, "agent_npn": "12345", "agent_name": "Agent A"}
        self.assertFalse(self.fc.check_agent_mismatch(record))

    def test_no_mismatch_when_both_resolve_same_uuid_case_insensitive(self):
        self.fc.resolve_agent_by_npn = lambda aid, npn, cid=None: ("UUID-ABC", "Resolved", [])
        self.fc.resolve_agent_by_name = lambda aid, name: ("uuid-abc", "Resolved", [])

        record = {"agency_id": 1, "agent_npn": "12345", "agent_name": "Agent A"}
        self.assertFalse(self.fc.check_agent_mismatch(record))

    def test_no_mismatch_when_name_unresolved(self):
        self.fc.resolve_agent_by_npn = lambda aid, npn, cid=None: ("agent-A-uuid", "Resolved", [])
        self.fc.resolve_agent_by_name = lambda aid, name: (None, "NotFound", [])

        record = {"agency_id": 1, "agent_npn": "12345", "agent_name": "Unregistered Agent"}
        self.assertFalse(self.fc.check_agent_mismatch(record))

    def test_no_mismatch_when_npn_unresolved(self):
        self.fc.resolve_agent_by_npn = lambda aid, npn, cid=None: (None, "NotFound", [])
        self.fc.resolve_agent_by_name = lambda aid, name: ("agent-B-uuid", "Resolved", [])

        record = {"agency_id": 1, "agent_npn": "UNKNOWN_ID", "agent_name": "Agent B"}
        self.assertFalse(self.fc.check_agent_mismatch(record))

    def test_no_mismatch_when_npn_is_ambiguous(self):
        self.fc.resolve_agent_by_npn = lambda aid, npn, cid=None: (None, "ConflictingMatches", [])
        self.fc.resolve_agent_by_name = lambda aid, name: ("agent-B-uuid", "Resolved", [])

        record = {"agency_id": 1, "agent_npn": "CONFLICT_ID", "agent_name": "Agent B"}
        self.assertFalse(self.fc.check_agent_mismatch(record))

    def test_no_mismatch_when_name_is_ambiguous(self):
        self.fc.resolve_agent_by_npn = lambda aid, npn, cid=None: ("agent-A-uuid", "Resolved", [])
        self.fc.resolve_agent_by_name = lambda aid, name: (None, "AmbiguousFullName", [])

        record = {"agency_id": 1, "agent_npn": "12345", "agent_name": "Common Name"}
        self.assertFalse(self.fc.check_agent_mismatch(record))

    def test_no_mismatch_when_missing_npn_or_name(self):
        self.fc.resolve_agent_by_npn = lambda aid, npn, cid=None: ("agent-A-uuid", "Resolved", [])
        self.fc.resolve_agent_by_name = lambda aid, name: ("agent-B-uuid", "Resolved", [])

        # Missing NPN
        self.assertFalse(self.fc.check_agent_mismatch({"agency_id": 1, "agent_npn": None, "agent_name": "Agent B"}))
        self.assertFalse(self.fc.check_agent_mismatch({"agency_id": 1, "agent_npn": "", "agent_name": "Agent B"}))
        # Missing Name
        self.assertFalse(self.fc.check_agent_mismatch({"agency_id": 1, "agent_npn": "12345", "agent_name": None}))
        self.assertFalse(self.fc.check_agent_mismatch({"agency_id": 1, "agent_npn": "12345", "agent_name": "   "}))
        # Both missing
        self.assertFalse(self.fc.check_agent_mismatch({"agency_id": 1, "agent_npn": None, "agent_name": None}))


class TestReconciliationAgentMismatch(unittest.TestCase):
    """Tests reconciliation behavior mirroring ValidateAgent.kt and ValidateAgentTest.kt."""

    def setUp(self):
        self.fc = forecast.Forecast.__new__(forecast.Forecast)
        self.fc.exceptions = []
        self.fc.local_policies = []
        self.fc.local_commissions = []
        self.fc.reattached_policies = set()
        self.fc.cancellation_gate = lambda record: False
        self.fc.maybe_reattach_commissions = lambda record: None
        self.fc.gross_has_value = lambda gross: False

    def test_reconciliation_fails_when_npn_and_name_resolve_to_different_agents(self):
        # Direct parity with ValidateAgentTest:
        # `reconciliation fails when Agent ID and Agent Name resolve to different agents`
        record = {
            "uuid": "rec-1",
            "policy_number": "POL-100",
            "policy_exists": True,
            "policy_status_id": POLICY_ACTIVE,
            "agency_id": 1,
            "agent_npn": "CA-123",
            "agent_name": "Agent B",
            "carrier_agent_appointment_number": "CA-123",  # Even if NPN matches policy appointment!
            "policy_uuid": "pol-uuid-1",
        }

        # Mock database selects for NPN and Name
        def mock_select(sql, params):
            if sql == AGENT_RESOLVE_SQL:
                return [{"agent_uuid": "agent-A-uuid", "resolution_method": "carrier_appointment", "priority": 2}]
            if sql == AGENT_NAME_RESOLVE_SQL:
                return [{"agent_uuid": "agent-B-uuid", "resolution_method": "full_name"}]
            return []

        self.fc.select = mock_select

        self.fc.reconcile_record(record)

        # Mismatch guard triggers immediately before agent matching
        self.assertEqual(len(self.fc.exceptions), 1)
        exc = self.fc.exceptions[0]
        self.assertEqual(exc["code"], AGENT_MISMATCH_CODE)
        self.assertEqual(exc["message"], AGENT_MISMATCH_ID_NAME_MESSAGE)
        self.assertEqual(exc["message"], "Agent ID and Agent Name belong to different agents")
        self.assertNotIn("match_remarks", record)

    def test_reconciliation_passes_when_npn_and_name_resolve_to_same_agent(self):
        # Direct parity with ValidateAgentTest:
        # `reconciliation passes when Agent ID and Agent Name resolve to the same agent`
        record = {
            "uuid": "rec-2",
            "policy_number": "POL-100",
            "policy_exists": True,
            "policy_status_id": POLICY_ACTIVE,
            "agency_id": 1,
            "agent_npn": "NPN-123",
            "agent_name": "John Doe",
            "carrier_agent_npn": "NPN-123",
            "policy_uuid": "pol-uuid-1",
            "gross": 0,
        }

        def mock_select(sql, params):
            if sql == AGENT_RESOLVE_SQL:
                return [{"agent_uuid": "same-agent-uuid", "resolution_method": "npn", "priority": 1}]
            if sql == AGENT_NAME_RESOLVE_SQL:
                return [{"agent_uuid": "same-agent-uuid", "resolution_method": "full_name"}]
            return []

        self.fc.select = mock_select

        self.fc.reconcile_record(record)

        agent_mismatches = [e for e in self.fc.exceptions if e["code"] == AGENT_MISMATCH_CODE]
        self.assertEqual(len(agent_mismatches), 0)
        self.assertIn("mapped with npn", record.get("match_remarks", ""))

    def test_reconciliation_passes_when_agent_id_matches_and_name_unresolved(self):
        # Direct parity with ValidateAgentTest:
        # `reconciliation passes when Agent ID matches policy and name does not resolve`
        record = {
            "uuid": "rec-3",
            "policy_number": "POL-100",
            "policy_exists": True,
            "policy_status_id": POLICY_ACTIVE,
            "agency_id": 1,
            "agent_npn": "CA-123",
            "agent_name": "Unregistered Agent",
            "carrier_agent_appointment_number": "CA-123",
            "policy_uuid": "pol-uuid-1",
            "gross": 0,
        }

        def mock_select(sql, params):
            if sql == AGENT_RESOLVE_SQL:
                return [{"agent_uuid": "agent-A-uuid", "resolution_method": "carrier_appointment", "priority": 2}]
            if sql == AGENT_NAME_RESOLVE_SQL:
                return []  # Name not found
            return []

        self.fc.select = mock_select

        self.fc.reconcile_record(record)

        agent_mismatches = [e for e in self.fc.exceptions if e["code"] == AGENT_MISMATCH_CODE]
        self.assertEqual(len(agent_mismatches), 0)
        self.assertIn("mapped with carrier_specific_id", record.get("match_remarks", ""))

    def test_reconciliation_passes_when_name_resolves_and_id_unresolved(self):
        record = {
            "uuid": "rec-4",
            "policy_number": "POL-100",
            "policy_exists": True,
            "policy_status_id": POLICY_ACTIVE,
            "agency_id": 1,
            "agent_npn": "UNKNOWN_ID",
            "agent_name": "Blair Agent",
            "carrier_agent_full_name": "Agent Blair",
            "policy_uuid": "pol-uuid-1",
            "gross": 0,
        }

        def mock_select(sql, params):
            if sql == AGENT_RESOLVE_SQL:
                return []  # NPN not found
            if sql == AGENT_NAME_RESOLVE_SQL:
                return [{"agent_uuid": "agent-B-uuid", "resolution_method": "full_name"}]
            return []

        self.fc.select = mock_select

        self.fc.reconcile_record(record)

        agent_mismatches = [e for e in self.fc.exceptions if e["code"] == AGENT_MISMATCH_CODE]
        self.assertEqual(len(agent_mismatches), 0)
        self.assertIn("mapped with agent_name", record.get("match_remarks", ""))

    def test_reconciliation_passes_when_identifier_lookup_is_ambiguous(self):
        # Ambiguous resolution does not raise AGENT_MISMATCH conflict
        record = {
            "uuid": "rec-5",
            "policy_number": "POL-100",
            "policy_exists": True,
            "policy_status_id": POLICY_ACTIVE,
            "agency_id": 1,
            "agent_npn": "CONFLICT_ID",
            "agent_name": "John Doe",
            "carrier_agent_npn": "CONFLICT_ID",
            "policy_uuid": "pol-uuid-1",
            "gross": 0,
        }

        def mock_select(sql, params):
            if sql == AGENT_RESOLVE_SQL:
                # Ambiguous NPN across branches
                return [
                    {"agent_uuid": "agent-A-uuid", "resolution_method": "npn", "priority": 1},
                    {"agent_uuid": "agent-B-uuid", "resolution_method": "user_code", "priority": 3},
                ]
            if sql == AGENT_NAME_RESOLVE_SQL:
                # Ambiguous Name
                return [
                    {"agent_uuid": "agent-C-uuid", "resolution_method": "full_name"},
                    {"agent_uuid": "agent-D-uuid", "resolution_method": "full_name"},
                ]
            return []

        self.fc.select = mock_select

        self.fc.reconcile_record(record)

        # No AGENT_MISMATCH conflict raised; falls through to agent_matches which matches NPN
        agent_mismatches = [e for e in self.fc.exceptions if e["code"] == AGENT_MISMATCH_CODE]
        self.assertEqual(len(agent_mismatches), 0)

    def test_reconciliation_passes_when_name_missing_and_id_matches(self):
        # Direct parity with ValidateAgentTest:
        # `reconciliation passes when only Agent ID is provided and matches policy agent`
        record = {
            "uuid": "rec-6",
            "policy_number": "POL-100",
            "policy_exists": True,
            "policy_status_id": POLICY_ACTIVE,
            "agency_id": 1,
            "agent_npn": "CA-123",
            "agent_name": None,
            "carrier_agent_appointment_number": "CA-123",
            "policy_uuid": "pol-uuid-1",
            "gross": 0,
        }

        self.fc.select = lambda sql, params: []

        self.fc.reconcile_record(record)

        agent_mismatches = [e for e in self.fc.exceptions if e["code"] == AGENT_MISMATCH_CODE]
        self.assertEqual(len(agent_mismatches), 0)
        self.assertIn("mapped with carrier_specific_id", record.get("match_remarks", ""))


class TestCreateLocalPolicyAgentMismatch(unittest.TestCase):
    """Tests create_local_policy mismatch validation mirroring PolicyCreationDaoTest.kt."""

    def setUp(self):
        self.fc = forecast.Forecast.__new__(forecast.Forecast)
        self.fc.exceptions = []
        self.fc.find_local_policy = lambda r: None

    def test_policy_creation_fails_when_npn_and_name_mismatch(self):
        # Parity with PolicyCreationDaoTest:
        # `resolveAgentUserForPolicyCreation returns Mismatch when ID and Name resolve to different UUIDs`
        record = {
            "uuid": "rec-1",
            "file_uuid": "file-1",
            "agency_id": 1,
            "agent_npn": "12345",
            "agent_name": "Agent B",
            "carrier_uuid": "c-1",
        }

        def mock_select(sql, params):
            if sql == forecast.POLICY_CREATE_DETAILS_SQL:
                return [{
                    "product_name": "Dental",
                    "product_uuid": "prod-1",
                    "carrier_uuid": "c-1",
                    "effective_date": "2026-01-01",
                }]
            if sql == AGENT_RESOLVE_SQL:
                return [{"agent_uuid": "agent-A-uuid", "resolution_method": "npn", "priority": 1}]
            if sql == AGENT_NAME_RESOLVE_SQL:
                return [{"agent_uuid": "agent-B-uuid", "resolution_method": "full_name"}]
            return []

        self.fc.select = mock_select

        res = self.fc.create_local_policy(record)

        self.assertFalse(res)
        self.assertEqual(len(self.fc.exceptions), 1)
        self.assertEqual(self.fc.exceptions[0]["code"], AGENT_MISMATCH_CODE)
        self.assertEqual(self.fc.exceptions[0]["message"], AGENT_MISMATCH_ID_NAME_MESSAGE)

    def test_policy_creation_succeeds_when_npn_and_name_resolve_to_same_uuid(self):
        # Parity with PolicyCreationDaoTest:
        # `resolveAgentUserForPolicyCreation returns Resolved when ID and Name resolve to same UUID`
        record = {
            "uuid": "rec-2",
            "file_uuid": "file-1",
            "agency_id": 1,
            "agent_npn": "12345",
            "agent_name": "Agent A",
            "carrier_uuid": "c-1",
        }

        def mock_select(sql, params):
            if sql == forecast.POLICY_CREATE_DETAILS_SQL:
                return [{
                    "product_name": "Dental",
                    "product_uuid": "prod-1",
                    "carrier_uuid": "c-1",
                    "effective_date": None,  # Stops after agent resolution successfully resolves
                }]
            if sql == AGENT_RESOLVE_SQL:
                return [{"agent_uuid": "same-agent-uuid", "resolution_method": "npn", "priority": 1}]
            if sql == AGENT_NAME_RESOLVE_SQL:
                return [{"agent_uuid": "same-agent-uuid", "resolution_method": "full_name"}]
            return []

        self.fc.select = mock_select

        res = self.fc.create_local_policy(record)

        self.assertFalse(res)
        # Did not fail on agent mismatch; failed on effective_date
        self.assertEqual(len(self.fc.exceptions), 1)
        self.assertIn("reason=effective_date", self.fc.exceptions[0]["message"])


if __name__ == "__main__":
    unittest.main()