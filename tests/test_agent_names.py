#!/usr/bin/env python3
"""Tests for Issue 9: Agent Name Normalization — Strict Backend Parity.

Derived from AgentNamesTest.kt, ValidateAgentTest.kt, and PolicyCreationDao.kt
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
    AGENT_NAME_RESOLVE_SQL,
    AgentNames,
    agent_matches,
    agent_name_matches,
    normalize_agent_name,
)


class TestAgentNamesNormalize(unittest.TestCase):
    """Parity tests for AgentNames.normalize against AgentNamesTest.kt."""

    def test_normalize_collapses_case_and_extra_spaces(self):
        # Direct parity from AgentNamesTest.kt
        self.assertEqual(AgentNames.normalize("Blair  Agent"), "blair agent")
        self.assertIsNone(AgentNames.normalize("  "))
        self.assertIsNone(AgentNames.normalize(None))

    def test_normalize_whitespace_variations(self):
        self.assertEqual(AgentNames.normalize("   Blair   Agent   "), "blair agent")
        self.assertEqual(AgentNames.normalize("Blair\t\n Agent"), "blair agent")
        self.assertIsNone(AgentNames.normalize(""))
        self.assertIsNone(AgentNames.normalize("\t\n  "))

    def test_normalize_single_and_multi_word(self):
        self.assertEqual(AgentNames.normalize("Madonna"), "madonna")
        self.assertEqual(AgentNames.normalize("Mary   Ann   Smith"), "mary ann smith")

    def test_module_level_alias(self):
        self.assertEqual(normalize_agent_name("Blair  Agent"), "blair agent")
        self.assertIsNone(normalize_agent_name(None))


class TestAgentNamesMatches(unittest.TestCase):
    """Parity tests for AgentNames.matches against AgentNamesTest.kt."""

    def test_matches_ignores_case_extra_spaces_and_swapped_first_last(self):
        # Direct parity from AgentNamesTest.kt
        self.assertTrue(AgentNames.matches("Blair  Agent", "Agent Blair"))
        self.assertTrue(AgentNames.matches("blair agent", "Blair Agent"))
        self.assertFalse(AgentNames.matches("Blair Agent", "Someone Else"))
        self.assertFalse(AgentNames.matches("Mary Ann Smith", "Smith Mary Ann"))

    def test_matches_swapped_both_directions(self):
        self.assertTrue(AgentNames.matches("Agent Blair", "Blair  Agent"))
        self.assertTrue(AgentNames.matches("  Blair   Agent  ", "  Agent   Blair  "))

    def test_matches_null_and_blank_handling(self):
        self.assertFalse(AgentNames.matches(None, "Blair Agent"))
        self.assertFalse(AgentNames.matches("Blair Agent", None))
        self.assertFalse(AgentNames.matches(None, None))
        self.assertFalse(AgentNames.matches("", "Blair Agent"))
        self.assertFalse(AgentNames.matches("Blair Agent", ""))
        self.assertFalse(AgentNames.matches("   ", "Blair Agent"))

    def test_matches_single_word(self):
        self.assertTrue(AgentNames.matches("Blair", "blair"))
        self.assertTrue(AgentNames.matches("  Blair  ", "blair"))
        self.assertFalse(AgentNames.matches("Blair", "Agent"))

    def test_matches_three_or_more_words_exact_only(self):
        # 3 words match only in exact word order (with whitespace/case collapsed)
        self.assertTrue(AgentNames.matches("Mary  Ann  Smith", "mary ann smith"))
        self.assertFalse(AgentNames.matches("Mary Ann Smith", "Smith Mary Ann"))
        self.assertFalse(AgentNames.matches("Mary Ann Smith", "Ann Mary Smith"))
        self.assertFalse(AgentNames.matches("John A B Doe", "Doe B A John"))

    def test_module_level_alias(self):
        self.assertTrue(agent_name_matches("Blair  Agent", "Agent Blair"))
        self.assertFalse(agent_name_matches("Blair Agent", "Someone Else"))


class TestAgentMatchesRecord(unittest.TestCase):
    """Tests agent_matches(record) during reconciliation (ValidateAgent.kt parity)."""

    def test_carrier_agent_full_name_swapped_with_spaces(self):
        # ValidateAgentTest.kt parity: reconciliation matches swapped agent name with extra spaces
        record = {
            "agent_npn": None,
            "agent_name": "Blair  Agent",
            "carrier_agent_appointment_number": None,
            "carrier_agent_full_name": "Agent Blair",
        }
        matched, remark = agent_matches(record)
        self.assertTrue(matched)
        self.assertEqual(remark, "mapped with agent_name")

    def test_writing_agent_full_name_swapped_with_spaces(self):
        record = {
            "agent_npn": None,
            "agent_name": "Blair  Agent",
            "writing_agent_full_name": "Agent Blair",
        }
        matched, remark = agent_matches(record)
        self.assertTrue(matched)
        self.assertEqual(remark, "mapped with agent_name")

    def test_exact_name_match_case_insensitive(self):
        record = {
            "agent_npn": None,
            "agent_name": "blair agent",
            "carrier_agent_full_name": "Blair Agent",
        }
        matched, remark = agent_matches(record)
        self.assertTrue(matched)
        self.assertEqual(remark, "mapped with agent_name")

    def test_different_name_does_not_match(self):
        record = {
            "agent_npn": None,
            "agent_name": "Blair Agent",
            "carrier_agent_full_name": "Someone Else",
            "writing_agent_full_name": "Other Person",
        }
        matched, remark = agent_matches(record)
        self.assertFalse(matched)
        self.assertIsNone(remark)

    def test_three_word_swapped_does_not_match(self):
        record = {
            "agent_npn": None,
            "agent_name": "Mary Ann Smith",
            "writing_agent_full_name": "Smith Mary Ann",
        }
        matched, remark = agent_matches(record)
        self.assertFalse(matched)
        self.assertIsNone(remark)

    def test_null_and_blank_agent_name(self):
        for blank_val in [None, "", "   "]:
            record = {
                "agent_npn": None,
                "agent_name": blank_val,
                "carrier_agent_full_name": "Agent Blair",
            }
            matched, remark = agent_matches(record)
            self.assertFalse(matched)
            self.assertIsNone(remark)

    def test_npn_matching_regression(self):
        # Existing NPN matching must continue to work and take priority
        record = {
            "agent_npn": "12345",
            "carrier_agent_npn": "12345",
            "carrier_agent_full_name": "Different Name",
        }
        matched, remark = agent_matches(record)
        self.assertTrue(matched)
        self.assertEqual(remark, "mapped with npn")


class TestAgentNameResolveSqlAndPolicyCreation(unittest.TestCase):
    """Tests AGENT_NAME_RESOLVE_SQL and create_local_policy (PolicyCreationDao.kt parity)."""

    def test_sql_structure_parity(self):
        # Must query contacts for entity = 'agent' and support both first/last and last/first
        sql_lower = AGENT_NAME_RESOLVE_SQL.lower()
        self.assertIn("from contacts c", sql_lower)
        self.assertIn("c.entity = 'agent'", sql_lower)
        self.assertIn("c.first_name || ' ' || c.last_name", sql_lower)
        self.assertIn("c.last_name || ' ' || c.first_name", sql_lower)
        self.assertIn("regexp_replace", sql_lower)

    def test_create_local_policy_normalizes_and_resolves_unique_agent(self):
        fc = forecast.Forecast.__new__(forecast.Forecast)
        fc.exceptions = []

        record = {
            "uuid": "rec-1",
            "file_uuid": "file-1",
            "agency_id": 5,
            "agent_npn": None,
            "agent_name": "Blair   Agent",
            "carrier_uuid": "c-1",
        }

        resolved_calls = []

        def mock_select(sql, params):
            if sql == forecast.POLICY_CREATE_DETAILS_SQL:
                return [{
                    "product_name": "Dental",
                    "product_uuid": "prod-1",
                    "carrier_uuid": "c-1",
                    "effective_date": None,  # Stops after agent resolution
                }]
            if sql == AGENT_NAME_RESOLVE_SQL:
                resolved_calls.append((sql, params))
                return [{"agent_uuid": "user-uuid-123", "resolution_method": "full_name"}]
            return []

        fc.select = mock_select
        fc.find_local_policy = lambda r: None

        res = fc.create_local_policy(record)
        # It stopped at missing effective_date after successfully resolving the agent
        self.assertFalse(res)
        self.assertEqual(len(resolved_calls), 1)
        # Confirms both %s parameters are the normalized agent name
        self.assertEqual(resolved_calls[0][1], (5, "blair agent", "blair agent"))
        # Confirms failure reason was NOT agent
        self.assertEqual(len(fc.exceptions), 1)
        self.assertIn("reason=effective_date", fc.exceptions[0]["message"])

    def test_create_local_policy_ambiguous_matches(self):
        fc = forecast.Forecast.__new__(forecast.Forecast)
        fc.exceptions = []

        record = {
            "uuid": "rec-2",
            "file_uuid": "file-1",
            "agency_id": 5,
            "agent_npn": None,
            "agent_name": "John Doe",
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
            if sql == AGENT_NAME_RESOLVE_SQL:
                return [
                    {"agent_uuid": "user-uuid-1", "resolution_method": "full_name"},
                    {"agent_uuid": "user-uuid-2", "resolution_method": "full_name"},
                ]
            return []

        fc.select = mock_select
        fc.find_local_policy = lambda r: None

        res = fc.create_local_policy(record)
        self.assertFalse(res)
        self.assertEqual(len(fc.exceptions), 1)
        self.assertIn("matched 2 users", fc.exceptions[0]["message"])

    def test_create_local_policy_zero_matches(self):
        fc = forecast.Forecast.__new__(forecast.Forecast)
        fc.exceptions = []

        record = {
            "uuid": "rec-3",
            "file_uuid": "file-1",
            "agency_id": 5,
            "agent_npn": None,
            "agent_name": "Ghost Agent",
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
            if sql == AGENT_NAME_RESOLVE_SQL:
                return []
            return []

        fc.select = mock_select
        fc.find_local_policy = lambda r: None

        res = fc.create_local_policy(record)
        self.assertFalse(res)
        self.assertEqual(len(fc.exceptions), 1)
        self.assertIn("could not be resolved", fc.exceptions[0]["message"])

    def test_create_local_policy_blank_name_and_missing_npn(self):
        fc = forecast.Forecast.__new__(forecast.Forecast)
        fc.exceptions = []

        record = {
            "uuid": "rec-4",
            "file_uuid": "file-1",
            "agency_id": 5,
            "agent_npn": "",
            "agent_name": "   ",
            "carrier_uuid": "c-1",
        }

        fc.find_local_policy = lambda r: None
        fc.select = lambda sql, params: []

        res = fc.create_local_policy(record)
        self.assertFalse(res)
        self.assertEqual(len(fc.exceptions), 1)
        self.assertIn("agent NPN and agent name are missing", fc.exceptions[0]["message"])


if __name__ == "__main__":
    unittest.main()