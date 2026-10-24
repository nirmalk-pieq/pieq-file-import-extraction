#!/usr/bin/env python3
"""Forecast review-to-reconcile processing for one or more commission statements.

Reads the environment database with SELECT only. POSTs a password grant
only to the Keycloak token URL derived from AUTH_URL. Never POSTs or PUTs
to PIEQ_API_BASE_URL and never writes the environment database.

Usage:
  python3 review_to_reconcile_file_forecast.py <preprod|prod> <commission_statement_uuid> [<uuid> ...]
  python3 review_to_reconcile_file_forecast.py <preprod|prod> review --agency_id <agency_id>

Loads config_preprod or config_prod from this folder. Environment variables
already set in the shell take precedence.

Environment:
  PIEQ_API_BASE_URL
  AUTH_URL
  client_id
  USERNAME
  PASSWORD
  grant_type
  PIEQ_DB_HOST
  PIEQ_DB_PORT
  PIEQ_DB_NAME
  PIEQ_DB_USERNAME
  PIEQ_DB_PASSWORD
"""

from __future__ import annotations

import csv
import importlib.util
import json
import logging
import os
import re
import sys
import uuid
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from urllib import error, parse, request

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

STATUS_EXCLUDED = -1
STATUS_FAILED = 0
STATUS_MATCHED = 1
FILE_STATUS_CLEARED = 0
FILE_STATUS_WAITING = 1
POLICY_ACTIVE = 0
PAYOUT_FIXED = 1
PAYOUT_PERCENT = 2
PERCENT_TOLERANCE = Decimal("1.00")
FMV_TOLERANCE = Decimal("1.00")
FMV_PROCESSING_TYPE = 1
MONEY = Decimal("0.01")
ENTITY_AGENCY = "agency"
ENTITY_AGENT = "agent"
ENTITY_SUBAGENT = "subagent"
SPLIT_PARTICIPANTS = {ENTITY_AGENCY, ENTITY_AGENT, ENTITY_SUBAGENT}
TYPE_COMMISSION = "COMMISSION"
TYPE_CHARGEBACK = "CHARGEBACK"
TYPE_TRUE_UP = "TRUE_UP"
TYPE_ARF = "ARF"
TYPE_AGENCY_CREDIT = "AGENCY_CREDIT"
DEBIT_TYPES = {"CHARGEBACK", "ADVANCE_RECOVERY", "AGENCY_DEBIT"}
POLICY_STATUS_LABELS = {
    0: "Active",
    1: "Pending",
    2: "Cancelled",
    3: "Expired",
    4: "Suspended",
    5: "In Force",
    6: "Lapsed",
}


class ForecastError(RuntimeError):
    pass


class TokenRejected(ForecastError):
    def __init__(self, status, url, detail):
        self.status = status
        super().__init__(f"GET {url} failed with status {status}: {detail}")


logger = logging.getLogger("review_to_reconcile_file_forecast")
row_logger = logging.getLogger("review_to_reconcile_file_forecast.row")


def log_row(record, detail, *args):
    def shown(value):
        if value is None:
            return "null"
        if isinstance(value, Decimal):
            return format(value, "f")
        return str(value)

    row_logger.info(
        "row=%s policy=%s uuid=%s gross=%s premium=%s chargeback=%s " + detail,
        record.get("row_order"),
        record.get("policy_number"),
        record.get("uuid"),
        shown(record.get("gross")),
        shown(record.get("premium_amount")),
        shown(record.get("chargeback_amount")),
        *args,
    )

SPLIT_COLUMNS = (
    "commission_split",
    "gross_commission",
    "earned_commission",
    "agency_split",
    "agent_split",
    "sales_leader_split",
    "subagent_split",
    "mmp_amount",
    "arf_amount",
)


def configure_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
        stream=sys.stderr,
    )
    row_logger.propagate = False
    row_logger.setLevel(logging.INFO)


def attach_forecast_log(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(path, mode="w", encoding="utf-8")
    handler.setLevel(logging.INFO)
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
    logger.addHandler(handler)
    row_logger.addHandler(handler)
    logger.info("Writing row log to %s", path)
    return handler


def detach_forecast_log(handler):
    logger.removeHandler(handler)
    row_logger.removeHandler(handler)
    handler.close()


def api_get(url, agency_id, token):
    if not token:
        raise ForecastError("Missing access token for read-only policy API calls")
    req = request.Request(
        url,
        method="GET",
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "x-client-id": "pieq",
            "x-tenant-id": str(agency_id),
            "x-service-name": "pieq-bg-fileProcessing",
        },
    )
    logger.info("GET %s", url)
    try:
        with request.urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode())
    except error.HTTPError as exc:
        detail = exc.read().decode(errors="replace")[:500]
        if exc.code in (401, 403):
            raise TokenRejected(exc.code, url, detail) from exc
        raise ForecastError(f"GET {url} failed with status {exc.code}: {detail}") from exc


def as_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for parser in (
        lambda raw: datetime.fromisoformat(raw.replace("Z", "+00:00")).date(),
        lambda raw: datetime.strptime(raw[:10], "%Y-%m-%d").date(),
        lambda raw: datetime.strptime(raw, "%m/%d/%Y").date(),
    ):
        try:
            return parser(text)
        except ValueError:
            continue
    return None


def statement_paid_to_date(record):
    """File processing reads data.paid_to_date. The table column is used only when that JSON value is empty."""
    data = json_obj(record.get("data"))
    if isinstance(data, dict):
        paid = as_date(data.get("paid_to_date"))
        if paid is not None:
            return paid
    return as_date(record.get("paid_to_date"))


def open_month_to(value):
    """A setup toMonth of 0 means the slab has no end, same as file processing."""
    if value is None or str(value).strip() == "":
        return None
    month = int(value)
    if month == 0:
        return None
    return month


REATTACH_MONTH_FROM = 1
REATTACH_DEFAULT_MAX_MONTH_TO = 12
COMMISSION_TYPE_REGULAR = 1


def calendar_day(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return as_date(value)


def policy_month_index(paid, effective):
    """Calendar months from the effective month to the paid-to month. The effective month is 0."""
    paid_day = calendar_day(paid)
    effective_day = calendar_day(effective)
    if paid_day is None or effective_day is None or paid_day < effective_day:
        return None
    return (paid_day.year - effective_day.year) * 12 + (paid_day.month - effective_day.month)


def commission_type_code(value):
    if value is None or str(value).strip() == "":
        return COMMISSION_TYPE_REGULAR
    try:
        return int(value)
    except (TypeError, ValueError):
        text = str(value).strip().upper()
    if text == "OVERRIDE":
        return 4
    if text == "BONUS":
        return 3
    return COMMISSION_TYPE_REGULAR


def slab_type_code(row):
    if row.get("commission_type") is not None:
        return commission_type_code(row.get("commission_type"))
    return commission_type_code(row.get("commission_type_text"))


def ranges_overlap(from_a, to_a, from_b, to_b):
    """Inclusive month ranges overlap when both ends are present. An open end does not pair."""
    if from_a is None or to_a is None or from_b is None or to_b is None:
        return False
    return max(int(from_a), int(from_b)) <= min(int(to_a), int(to_b))


def plan_reattach_actions(existing, incoming):
    """Pair stored slabs with the latest product setup, using the same rules as file processing."""
    existing_by_type = {}
    incoming_by_type = {}
    for row in existing:
        existing_by_type.setdefault(slab_type_code(row), []).append(row)
    for row in incoming:
        incoming_by_type.setdefault(slab_type_code(row), []).append(row)
    actions = []
    for commission_type in sorted(set(existing_by_type) | set(incoming_by_type)):
        actions.extend(
            plan_reattach_type(existing_by_type.get(commission_type, []), incoming_by_type.get(commission_type, []))
        )
    return actions


def plan_reattach_type(existing, incoming):
    if not incoming:
        return []
    if any(not row.get("commission_setup_uuid") for row in incoming):
        return []
    if not existing:
        return [{"action": "insert", "incoming": row} for row in incoming]
    setup_ids = {str(row.get("commission_setup_uuid")) for row in incoming}
    if len(setup_ids) == 1 and all(str(row.get("commission_setup_uuid") or "") == next(iter(setup_ids)) for row in existing):
        return []
    existing_sorted = sorted(existing, key=lambda row: (row.get("month_from") if row.get("month_from") is not None else 10**9, str(row.get("uuid") or "")))
    incoming_sorted = sorted(incoming, key=lambda row: (row.get("month_from") if row.get("month_from") is not None else 10**9, str(row.get("commission_type_text") or "")))
    used = set()
    actions = []
    for row in existing_sorted:
        match_idx = next(
            (
                idx
                for idx, candidate in enumerate(incoming_sorted)
                if idx not in used
                and ranges_overlap(row.get("month_from"), row.get("month_to"), candidate.get("month_from"), candidate.get("month_to"))
            ),
            None,
        )
        if match_idx is None:
            actions.append({"action": "delete", "existing": row})
            continue
        used.add(match_idx)
        actions.append({"action": "update", "existing": row, "incoming": incoming_sorted[match_idx]})
    for idx, candidate in enumerate(incoming_sorted):
        if idx not in used:
            actions.append({"action": "insert", "incoming": candidate})
    return actions


def overlay_split_values(existing, incoming):
    existing_rows = existing if isinstance(existing, list) else json_obj(existing)
    incoming_rows = incoming if isinstance(incoming, list) else []
    if not isinstance(existing_rows, list) or not existing_rows:
        return list(incoming_rows)
    by_order = {}
    by_role = {}
    for item in incoming_rows:
        if not isinstance(item, dict):
            continue
        if item.get("order") is not None:
            by_order[int(item["order"])] = item
        by_role[str(item.get("role") or "").lower()] = item
    updated = []
    for item in existing_rows:
        if not isinstance(item, dict):
            continue
        copied = dict(item)
        source = by_order.get(int(item["order"])) if item.get("order") is not None else None
        if source is None:
            source = by_role.get(str(item.get("role") or "").lower())
        if source is not None and source.get("commission_value") is not None:
            copied["commission_value"] = source.get("commission_value")
        updated.append(copied)
    return updated


def apply_reattach_actions(existing, actions):
    removed = set()
    produced = []
    for action in actions:
        kind = action["action"]
        if kind == "delete":
            removed.add(action["existing"].get("uuid"))
            continue
        if kind == "update":
            current = action["existing"]
            incoming = action["incoming"]
            removed.add(current.get("uuid"))
            updated = dict(current)
            for field in (
                "fee_amount",
                "payout_method",
                "premium_amount",
                "month_from",
                "month_to",
                "commission_type",
                "commission_type_text",
                "commission_setup_uuid",
            ):
                updated[field] = incoming.get(field)
            updated["uuid"] = current.get("uuid")
            updated["commission_split"] = overlay_split_values(current.get("commission_split"), incoming.get("commission_split"))
            updated.pop("period", None)
            produced.append(updated)
            continue
        inserted = dict(action["incoming"])
        inserted["uuid"] = new_id()
        inserted.pop("period", None)
        produced.append(inserted)
    kept = [row for row in existing if row.get("uuid") not in removed]
    return kept + produced


def nested_amount(value):
    if isinstance(value, dict):
        return money(value.get("value") or value.get("commissionValue") or value.get("commission_value"))
    return money(value)


ROLE_AGENCY = "Agency"
ROLE_SALES_LEADER = "Sales Leader"
ROLE_AGENT = "Agent"
ROLE_SUB_AGENT = "Sub Agent"
ROLE_ORDER = (ROLE_AGENCY, ROLE_SALES_LEADER, ROLE_AGENT, ROLE_SUB_AGENT)


def is_subagent_entity(entity):
    return str(entity or "").strip().lower() == "subagent"


def is_subagent_role_name(name):
    return str(name or "").strip().lower() in ("sub agent", "sub-agent", "subagent")


def hierarchy_user_full_name(user):
    if not isinstance(user, dict):
        return ""
    parts = [user.get("firstName") or user.get("first_name"), user.get("lastName") or user.get("last_name")]
    return " ".join(part for part in parts if part and str(part).strip())


def hierarchy_level_uuid(user):
    if not isinstance(user, dict):
        return None
    level = user.get("level") or {}
    if isinstance(level, dict):
        return level.get("uuid")
    return None


def build_hierarchy_chain(hierarchy_users, pay_to_agent_uuid):
    by_uuid = {}
    for user in hierarchy_users or []:
        uuid = user.get("uuid") if isinstance(user, dict) else None
        if uuid:
            by_uuid[str(uuid)] = user
    chain = []
    current = str(pay_to_agent_uuid) if pay_to_agent_uuid else None
    while current and current in by_uuid:
        user = by_uuid[current]
        chain.append(user)
        current = user.get("reporting_manager") or user.get("reportingManager")
    return chain


def agency_hierarchy_levels(payload, hierarchy_agency_uuid):
    data = (payload or {}).get("data") or {}
    agencies = data.get("agencies") or []
    if agencies:
        want = str(hierarchy_agency_uuid or "").lower()
        for agency in agencies:
            if str(agency.get("uuid") or "").lower() == want:
                return agency.get("levels") or []
        return []
    return data.get("levels") or []


def build_level_order_maps(levels):
    by_order = {}
    by_order_agent = {}
    by_order_subagent = {}
    for level in levels or []:
        order = level.get("levelOrder")
        if order is None:
            order = level.get("level_order")
        uuid = level.get("uuid")
        if order is None or not uuid:
            continue
        order = int(order)
        by_order[order] = str(uuid)
        entity = str(level.get("entity") or "").strip().lower()
        if entity == "agent":
            by_order_agent[order] = str(uuid)
        elif entity == "subagent":
            by_order_subagent[order] = str(uuid)
    return {
        "flat": by_order,
        "agent": by_order_agent,
        "subagent": by_order_subagent,
    }


def _level_order(wire_level):
    order = wire_level.get("level_order")
    if order is None:
        order = wire_level.get("levelOrder")
    return int(order) if order is not None else None


def _wire_commission_value(wire_level):
    return money(wire_level.get("commission_value") if "commission_value" in wire_level else wire_level.get("commissionValue")) or Decimal("0")


def normalize_match_role(role, maps):
    role_name = role.get("name")
    normalized = []
    for wire in role.get("levels") or []:
        order = _level_order(wire)
        if order is None:
            continue
        if is_subagent_role_name(role_name):
            tier_uuid = maps["subagent"].get(order) or maps["flat"].get(order)
            if not tier_uuid:
                continue
            children = []
            child_block = wire.get("children") or {}
            for child in child_block.get("levels") or []:
                child_order = _level_order(child)
                if child_order is None:
                    continue
                child_uuid = maps["agent"].get(child_order)
                if not child_uuid:
                    continue
                children.append(
                    {
                        "levelUuid": child_uuid,
                        "commissionValue": {"value": _wire_commission_value(child)},
                    }
                )
            children.sort(key=lambda item: item.get("levelUuid") or "")
            normalized.append(
                {
                    "levelUuid": tier_uuid,
                    "commissionValue": {"value": _wire_commission_value(wire)},
                    "children": {"levels": children} if children else None,
                }
            )
        else:
            level_uuid = maps["flat"].get(order)
            if not level_uuid:
                continue
            normalized.append(
                {
                    "levelUuid": level_uuid,
                    "commissionValue": {"value": _wire_commission_value(wire)},
                    "children": None,
                }
            )
    return {"name": role_name, "levels": normalized}


def select_commission_group(period, affiliated_agency_uuid, fallback_agency_uuid):
    groups = period.get("commissionSplits") or period.get("commission_splits") or []
    if not groups:
        return None, fallback_agency_uuid
    if affiliated_agency_uuid:
        want = str(affiliated_agency_uuid).lower()
        for group in groups:
            group_agency = group.get("commissionAgencyUuid") or group.get("commission_agency_uuid")
            if group_agency and str(group_agency).lower() == want:
                return group, str(group_agency)
    group = groups[0]
    group_agency = group.get("commissionAgencyUuid") or group.get("commission_agency_uuid") or fallback_agency_uuid
    return group, str(group_agency) if group_agency else fallback_agency_uuid


def level_commission_value(levels, level_uuid):
    if not levels:
        return Decimal("0")
    if level_uuid:
        for level in levels:
            if str(level.get("levelUuid") or "") == str(level_uuid):
                return nested_amount(level.get("commissionValue")) or Decimal("0")
        return Decimal("0")
    return nested_amount(levels[0].get("commissionValue")) or Decimal("0")


def find_subagent_nested_percent(normalized_roles, pay_to_tier_uuid, agent_grade_uuid):
    sub_role = None
    for role in normalized_roles:
        if is_subagent_role_name(role.get("name")):
            sub_role = role
            break
    if not sub_role:
        return None
    tier = None
    for level in sub_role.get("levels") or []:
        if str(level.get("levelUuid") or "") == str(pay_to_tier_uuid):
            tier = level
            break
    if not tier:
        return None
    for nested in (tier.get("children") or {}).get("levels") or []:
        if str(nested.get("levelUuid") or "") == str(agent_grade_uuid):
            return nested_amount(nested.get("commissionValue"))
    return None


def build_commission_splits_for_period(
    period,
    pay_to_agent_uuid,
    tenant_agency_uuid,
    hierarchy_users,
    level_maps,
):
    """Build commission_split lines using agent grade / reporting chain like fileprocessing attach."""
    chain = build_hierarchy_chain(hierarchy_users, pay_to_agent_uuid)
    if not chain:
        raise ForecastError(
            f"Reporting hierarchy data missing pay-to agent {pay_to_agent_uuid}"
        )
    pay_to = chain[0]
    pay_to_level_uuid = hierarchy_level_uuid(pay_to)
    pay_to_sub_agent = is_subagent_entity(pay_to.get("entity"))
    manager_level_uuid = hierarchy_level_uuid(chain[1]) if len(chain) > 1 else None
    if pay_to_sub_agent and len(chain) < 2:
        raise ForecastError(
            f"Sub-Agent commission attach skipped: reporting chain length {len(chain)} < 2"
        )

    affiliated = pay_to.get("affiliatedAgencyUuid") or pay_to.get("affiliated_agency_uuid")
    group, group_agency = select_commission_group(period, affiliated, tenant_agency_uuid)
    raw_roles = (group or {}).get("roles") or []
    normalized_roles = [normalize_match_role(role, level_maps) for role in raw_roles]

    level_uuid_for_flat = manager_level_uuid if pay_to_sub_agent else pay_to_level_uuid
    if pay_to_sub_agent and not level_uuid_for_flat:
        level_uuid_for_flat = pay_to_level_uuid

    amounts = {role: Decimal("0") for role in ROLE_ORDER}
    for role in normalized_roles:
        role_name = role.get("name") or ""
        if is_subagent_role_name(role_name):
            continue
        for canonical in ROLE_ORDER:
            if role_name.strip().lower() == canonical.lower():
                amounts[canonical] = level_commission_value(role.get("levels") or [], level_uuid_for_flat)
                break

    if pay_to_sub_agent and pay_to_level_uuid and manager_level_uuid:
        agent_raw = amounts[ROLE_AGENT]
        if agent_raw > 0:
            percent = find_subagent_nested_percent(normalized_roles, pay_to_level_uuid, manager_level_uuid)
            if percent is not None:
                sub_amount = (percent * agent_raw / Decimal("100")).quantize(MONEY, rounding=ROUND_HALF_UP)
                amounts[ROLE_SUB_AGENT] = sub_amount
                amounts[ROLE_AGENT] = (agent_raw - sub_amount).quantize(MONEY, rounding=ROUND_HALF_UP)
                if amounts[ROLE_AGENT] < 0:
                    amounts[ROLE_AGENT] = Decimal("0.00")

    chain_length = len(chain)

    def seat(user):
        if not user:
            return None, None
        return str(user.get("uuid")), str(user.get("entity") or "agent")

    if chain_length >= 3:
        sub_uuid, sub_entity = seat(chain[0])
        agent_uuid, agent_entity = seat(chain[1])
        leader_uuid, leader_entity = seat(chain[2])
        sub_amount = amounts[ROLE_SUB_AGENT]
    elif chain_length == 2 and pay_to_sub_agent:
        sub_uuid, sub_entity = seat(chain[0])
        agent_uuid, agent_entity = seat(chain[1])
        leader_uuid, leader_entity = agent_uuid, agent_entity
        sub_amount = amounts[ROLE_SUB_AGENT]
    elif chain_length == 2:
        sub_uuid, sub_entity, sub_amount = None, None, Decimal("0")
        agent_uuid, agent_entity = seat(chain[0])
        leader_uuid, leader_entity = seat(chain[1])
    else:
        sub_uuid, sub_entity, sub_amount = None, None, Decimal("0")
        agent_uuid, agent_entity = seat(chain[0])
        leader_uuid, leader_entity = agent_uuid, agent_entity

    mapping = [
        (1, ROLE_AGENCY, "agency", group_agency, amounts[ROLE_AGENCY]),
        (2, ROLE_SALES_LEADER, leader_entity or "agent", leader_uuid, amounts[ROLE_SALES_LEADER]),
        (3, ROLE_AGENT, agent_entity or "agent", agent_uuid, amounts[ROLE_AGENT]),
        (4, ROLE_SUB_AGENT, sub_entity or "subagent", sub_uuid, sub_amount),
    ]
    lines = []
    for order, role_name, entity, user_uuid, value in mapping:
        if not user_uuid:
            continue
        lines.append(
            {
                "order": order,
                "role": role_name,
                "entity": entity,
                "user_uuid": str(user_uuid),
                "commission_value": value,
            }
        )
    return lines


def slabs_from_match_response(payload, policy_uuid):
    slabs = []
    for setup in payload.get("data") or []:
        inner = setup.get("data") or {}
        status = str(inner.get("status") or "published").lower()
        if status not in ("published", ""):
            continue
        ipv_ranges = inner.get("ipvRanges") or inner.get("ipv_ranges") or []
        for ipv in ipv_ranges:
            for period in ipv.get("periods") or []:
                earning = str(period.get("earningType") or period.get("earning_type") or "").lower()
                month_from = period.get("fromMonth")
                if month_from is None:
                    month_from = period.get("from_month")
                month_to = period.get("toMonth") if "toMonth" in period else period.get("to_month")
                type_text = setup.get("type") or "COMMISSION"
                slabs.append(
                    {
                        "uuid": new_id(),
                        "policy_uuid": policy_uuid,
                        "fee_amount": money(period.get("value")),
                        "payout_method": PAYOUT_PERCENT if earning == "percentage" else PAYOUT_FIXED,
                        "premium_amount": None,
                        "commission_type": commission_type_code(type_text),
                        "commission_type_text": type_text,
                        "month_from": int(month_from) if month_from is not None else 1,
                        "month_to": open_month_to(month_to),
                        "commission_setup_uuid": setup.get("commissionSetupUuid") or setup.get("uuid"),
                        "period": period,
                        "commission_split": [],
                    }
                )
    return slabs


POLICY_CREATE_DETAILS_SQL = """
SELECT
    csr.policy_number,
    (NULLIF(TRIM(csr.data->>'policy_effective_date'), ''))::timestamp AS effective_date,
    csr.data->>'product_name' AS product_name,
    NULLIF(TRIM(SPLIT_PART(csr.data->>'policy_holder_name', ',', 1)), '') AS holder_last_name,
    NULLIF(TRIM(SPLIT_PART(csr.data->>'policy_holder_name', ',', 2)), '') AS holder_first_name,
    cs.carrier_uuid::text AS carrier_uuid,
    pa.product_uuid::text AS product_uuid
FROM commission_statement_records csr
INNER JOIN commission_statements cs
    ON cs.agency_id = csr.agency_id AND cs.uuid = csr.file_uuid
LEFT JOIN product_aliases pa
    ON pa.agency_id = csr.agency_id
   AND pa.alias = csr.data->>'product_name'
   AND pa.carrier_id = cs.carrier_uuid
WHERE csr.agency_id = %s AND csr.uuid = %s::uuid AND csr.file_uuid = %s::uuid
"""

AGENT_RESOLVE_SQL = """
SELECT agent_uuid::text AS agent_uuid, resolution_method, priority
FROM (
    SELECT a.user_uuid AS agent_uuid, 'npn' AS resolution_method, 1 AS priority
    FROM agents a
    WHERE a.agency_id = %s AND a.id > 0 AND a.npn = %s
    UNION ALL
    SELECT ca.agent_id AS agent_uuid, 'carrier_appointment' AS resolution_method, 2 AS priority
    FROM carrier_appointments ca
    WHERE ca.agency_id = %s AND ca.id > 0 AND ca.number = %s
      AND ca.carrier_id = %s::uuid
    UNION ALL
    SELECT u.uuid AS agent_uuid, 'user_code' AS resolution_method, 3 AS priority
    FROM users u
    WHERE u.tenant_id = %s AND u.id > 0 AND u.code = %s
) combined
WHERE agent_uuid IS NOT NULL
ORDER BY priority
"""

AGENT_NAME_RESOLVE_SQL = """
SELECT DISTINCT c.relationship_id::text AS agent_uuid, 'full_name' AS resolution_method
FROM contacts c
WHERE c.agency_id = %s
  AND c.id > 0
  AND c.entity = 'agent'
  AND (
      regexp_replace(lower(trim(c.first_name || ' ' || c.last_name)), '\\s+', ' ', 'g') = %s
      OR regexp_replace(lower(trim(c.last_name || ' ' || c.first_name)), '\\s+', ' ', 'g') = %s
  )
"""

AGENT_MISMATCH_CODE = "AGENT_MISMATCH"
AGENT_MISMATCH_ID_NAME_MESSAGE = "Agent ID and Agent Name belong to different agents"



REPEATED_WHITESPACE = re.compile(r"\s+")


class AgentNames:
    """Normalizes a statement agent name before a contact lookup.

    Case and repeated spaces are ignored. A two-word name matches in either order.
    A name with more than two words matches only as written.
    """

    @staticmethod
    def normalize(name: str | None) -> str | None:
        if name is None:
            return None
        text = REPEATED_WHITESPACE.sub(" ", str(name).strip().lower())
        return text if text else None

    @classmethod
    def matches(cls, file_name: str | None, contact_full_name: str | None) -> bool:
        file = cls.normalize(file_name)
        if not file:
            return False
        contact = cls.normalize(contact_full_name)
        if not contact:
            return False
        if file == contact:
            return True
        swapped = cls.swap_two_word_name(file)
        if not swapped:
            return False
        return swapped == contact

    @staticmethod
    def swap_two_word_name(name: str) -> str | None:
        parts = name.split(" ")
        if len(parts) != 2:
            return None
        return f"{parts[1]} {parts[0]}"


normalize_agent_name = AgentNames.normalize
agent_name_matches = AgentNames.matches


def money(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1].strip()
    text = text.replace(",", "").replace("$", "").strip()
    try:
        amount = Decimal(text).quantize(MONEY, rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return None
    return -amount if negative else amount


def as_bool(value):
    return bool(value) if value is not None else False


def json_obj(value):
    if value is None:
        return {}
    if isinstance(value, (dict, list)):
        return value
    if isinstance(value, str):
        text = value.strip()
        return json.loads(text) if text else {}
    return value


def new_id():
    return str(uuid.uuid4())


def load_env_file(environment):
    env_path = Path(__file__).resolve().parent / f"config_{environment}"
    if not env_path.exists():
        raise ForecastError(f"Config file not found: {env_path}")
    loaded = 0
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value
            loaded += 1
    logger.info("Loaded %s value(s) from %s", loaded, env_path.name)


def require_env(name):
    value = os.environ.get(name, "").strip()
    if not value:
        raise ForecastError(f"Missing environment variable {name}")
    return value


def normalize_db_url(raw):
    url = raw.strip()
    return url[len("jdbc:") :] if url.startswith("jdbc:") else url


def resolve_database():
    password = require_env("PIEQ_DB_PASSWORD")
    explicit = os.environ.get("PIEQ_DB_CONNECTION_STRING", "").strip()
    if explicit and "://" in explicit:
        return explicit, password
    host = explicit or os.environ.get("PIEQ_DB_HOST", "").strip()
    if not host:
        raise ForecastError("Set PIEQ_DB_HOST or PIEQ_DB_CONNECTION_STRING")
    user = require_env("PIEQ_DB_USERNAME")
    database = require_env("PIEQ_DB_NAME")
    port = os.environ.get("PIEQ_DB_PORT", "5432").strip() or "5432"
    url = f"postgresql://{parse.quote(user, safe='')}@{host}:{port}/{database}"
    return url, password


def db_target_label(url):
    parsed = parse.urlparse(normalize_db_url(url))
    host = parsed.hostname or "unknown-host"
    port = parsed.port or 5432
    database = (parsed.path or "").lstrip("/") or "unknown-db"
    user = parsed.username or "password-from-env"
    return f"host={host} port={port} db={database} user={user}"


def connect_readonly(url, password):
    conninfo = normalize_db_url(url)
    label = db_target_label(url)
    logger.info("Connecting to database (%s), read only", label)
    conn = None
    driver = None
    try:
        try:
            import psycopg

            conn = psycopg.connect(conninfo, password=password, autocommit=True)
            driver = "psycopg"
        except ImportError:
            import psycopg2

            conn = psycopg2.connect(conninfo, password=password)
            conn.set_session(readonly=True, autocommit=True)
            driver = "psycopg2"
    except ImportError as exc:
        raise ForecastError("Install psycopg or psycopg2 to run this script") from exc
    except Exception as exc:
        logger.error("Database connection failed (%s): %s", label, exc)
        raise ForecastError(f"Database connection failed ({label}): {exc}") from exc

    try:
        if driver == "psycopg":
            conn.execute("SET default_transaction_read_only = on")
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT current_database(), current_user,
                       COALESCE(inet_server_addr()::text, 'local'),
                       COALESCE(inet_server_port()::text, '')
                """
            )
            database, db_user, server_addr, server_port = cur.fetchone()
    except Exception as exc:
        conn.close()
        logger.error("Database connection failed after login (%s): %s", label, exc)
        raise ForecastError(f"Database connection failed ({label}): {exc}") from exc

    server = server_addr if not server_port else f"{server_addr}:{server_port}"
    logger.info(
        "Database connected: db=%s user=%s server=%s driver=%s read_only=on",
        database,
        db_user,
        server,
        driver,
    )
    return conn


def fetch_all(conn, sql, params):
    stripped = sql.lstrip().lower()
    if not (stripped.startswith("select") or stripped.startswith("with")):
        raise ForecastError("Refusing to run a statement that is not SELECT")
    with conn.cursor() as cur:
        cur.execute(sql, params)
        columns = [col.name for col in cur.description]
        return [dict(zip(columns, row)) for row in cur.fetchall()]


def sso_token_url(auth_url):
    raw = auth_url.strip().rstrip("/")
    if raw.endswith("/token"):
        return raw
    return f"{raw}/realms/pieq-sso/protocol/openid-connect/token"


def obtain_token():
    auth_url = (os.environ.get("AUTH_URL") or os.environ.get("PIEQ_TOKEN_ENDPOINT") or "").strip()
    if not auth_url:
        raise ForecastError("Missing environment variable AUTH_URL")
    token_url = sso_token_url(auth_url)
    base_url = os.environ.get("PIEQ_API_BASE_URL", "").strip()
    if base_url and token_url.rstrip("/") == base_url.rstrip("/"):
        raise ForecastError("Token URL must not be the API base URL")
    client_id = (os.environ.get("client_id") or os.environ.get("KEYCLOAK_CLIENT_ID") or "pieq-app").strip()
    username = (os.environ.get("USERNAME") or os.environ.get("PIEQ_API_USERNAME") or "").strip()
    password = (os.environ.get("PASSWORD") or os.environ.get("PIEQ_API_PASSWORD") or "").strip()
    grant_type = (os.environ.get("grant_type") or "password").strip()
    if not username or not password:
        raise ForecastError("Missing USERNAME or PASSWORD")
    body = parse.urlencode(
        {
            "grant_type": grant_type,
            "client_id": client_id,
            "username": username,
            "password": password,
        }
    ).encode()
    req = request.Request(
        token_url,
        data=body,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"},
    )
    logger.info("Requesting token for %s from %s (client_id=%s, grant_type=%s)", username, token_url, client_id, grant_type)
    try:
        with request.urlopen(req, timeout=60) as resp:
            payload = json.loads(resp.read().decode())
    except error.HTTPError as exc:
        detail = exc.read().decode(errors="replace")[:500]
        raise ForecastError(f"Token request failed with status {exc.code}: {detail}") from exc
    token = payload.get("access_token") or payload.get("token")
    if not token:
        raise ForecastError("Token response did not contain access_token")
    logger.info("Token obtained")
    return token


def _json_default(value):
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


def sql_literal(value):
    if value is None:
        return "NULL"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, (datetime, date)):
        return "'" + value.isoformat().replace("'", "''") + "'"
    if isinstance(value, (dict, list)):
        return "'" + json.dumps(value, default=_json_default).replace("'", "''") + "'::jsonb"
    return "'" + str(value).replace("'", "''") + "'"


def insert_sql(table, row):
    columns = list(row.keys())
    values = ", ".join(sql_literal(row[column]) for column in columns)
    return f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({values});"


def effective_members(total_members) -> int:
    if total_members is None:
        return 1
    count = int(Decimal(str(total_members)))
    return 1 if count <= 1 else count


RECORD_SQL = """
SELECT DISTINCT ON (csr.id)
    csr.id, csr.uuid::text AS uuid, csr.agency_id, csr.file_uuid::text AS file_uuid,
    csr.row_order, csr.policy_number,
    NULLIF(csr.data->>'agent_npn', '') AS agent_npn,
    NULLIF(csr.data->>'agent_name', '') AS agent_name,
    NULLIF(TRIM(csr.data->>'gross_commission'), '') AS gross_commission,
    NULLIF(TRIM(csr.data->>'premium'), '') AS premium,
    NULLIF(TRIM(COALESCE(csr.data->>'commission_percent', csr.data->>'commission_percentage')), '') AS commission_percent,
    NULLIF(TRIM(csr.data->>'chargeback'), '') AS chargeback,
    NULLIF(TRIM(csr.data->>'earned_commission'), '') AS earned_commission,
    NULLIF(TRIM(csr.data->>'total_members'), '') AS total_members,
    csr.data AS data, csr.paid_to_date, csr.policy_uuid::text AS stored_policy_uuid,
    csr.policy_commission_uuid::text AS policy_commission_uuid, csr.commission_type,
    csr.chargeback_uuid::text AS chargeback_uuid, csr.chargeback_entity_type,
    csr.is_commission_validated, csr.is_chargeback_validated,
    csr.is_policy_cancellation_validated, csr.is_advance_validated,
    csr.is_true_up_validated, csr.true_up_adjustment, csr.status, csr.match_remarks, csr.error,
    cs.name AS source_file_name, cs.batch_id AS source_batch_id,
    cs.updated_at AS statement_updated_at, cs.created_at AS statement_created_at,
    cs.carrier_uuid::text AS carrier_uuid,
    p.uuid::text AS policy_uuid, p.status_id AS policy_status_id,
    p.effective_date AS policy_effective_date, p.carrier_id::text AS policy_carrier_id,
    p.product_id::text AS policy_product_uuid, p.pay_to_agent_user_id::text AS pay_to_agent_user_id,
    (p.id IS NOT NULL) AS policy_exists,
    pa.initiated_by AS policy_advance_initiated_by,
    pa.process_advance AS policy_advance_process_advance,
    pa.process_commission AS policy_advance_process_commission,
    pa.entity_type AS policy_advance_entity_type, pa.entity_uuid::text AS policy_advance_entity_uuid,
    pa.status AS policy_advance_status, pa.advance_amount AS policy_advance_amount,
    pa.recovered_amount AS policy_advance_recovered_amount,
    (pa.uuid IS NOT NULL) AS has_policy_advance,
    csr.transaction_type,
    prod_fmv.commission_processing_type AS commission_processing_type,
    u_carrier.code AS carrier_agent_user_code,
    TRIM(CONCAT(contacts_carrier.first_name, ' ', contacts_carrier.last_name)) AS carrier_agent_full_name,
    a_carrier.npn AS carrier_agent_npn, ca_carrier.number AS carrier_agent_appointment_number,
    u_writing.code AS writing_agent_user_code,
    TRIM(CONCAT(contacts_writing.first_name, ' ', contacts_writing.last_name)) AS writing_agent_full_name,
    a_writing.npn AS writing_agent_npn, ca_writing.number AS writing_agent_appointment_number
FROM commission_statement_records csr
JOIN commission_statements cs ON cs.uuid = csr.file_uuid AND cs.agency_id = csr.agency_id
LEFT JOIN policies p
  ON p.agency_id = csr.agency_id AND p.number = csr.policy_number AND p.carrier_id = cs.carrier_uuid
LEFT JOIN policy_advances pa ON pa.agency_id = p.agency_id AND pa.policy_uuid = p.uuid
LEFT JOIN users u_carrier ON u_carrier.uuid = p.carrier_agent_user_id
LEFT JOIN contacts contacts_carrier
  ON contacts_carrier.relationship_id = u_carrier.uuid AND contacts_carrier.agency_id = csr.agency_id
LEFT JOIN agents a_carrier ON a_carrier.user_uuid = u_carrier.uuid
LEFT JOIN carrier_appointments ca_carrier
  ON ca_carrier.agent_id = u_carrier.uuid AND ca_carrier.carrier_id = cs.carrier_uuid
LEFT JOIN users u_writing ON u_writing.uuid = p.writing_agent_user_id
LEFT JOIN contacts contacts_writing
  ON contacts_writing.relationship_id = u_writing.uuid AND contacts_writing.agency_id = csr.agency_id
LEFT JOIN agents a_writing ON a_writing.user_uuid = u_writing.uuid
LEFT JOIN carrier_appointments ca_writing
  ON ca_writing.agent_id = u_writing.uuid AND ca_writing.carrier_id = cs.carrier_uuid
LEFT JOIN products prod_fmv
  ON prod_fmv.agency_id = p.agency_id AND prod_fmv.uuid = p.product_id
WHERE csr.file_uuid = %s::uuid
ORDER BY csr.id, csr.row_order
"""

# Calendar month difference from effective date to paid-to date. The effective month is 0.
# February paid-to on a January effective date is month 1. Same formula as file processing.
POLICY_MONTH_EXPR = """(
    (EXTRACT(YEAR FROM %s::date) - EXTRACT(YEAR FROM %s::date)) * 12
    + (EXTRACT(MONTH FROM %s::date) - EXTRACT(MONTH FROM %s::date))
)"""

MATCH_SQL = f"""
SELECT pc.uuid::text AS uuid, pc.fee_amount, pc.payout_method, pc.premium_amount,
       pc.commission_type_text, pc.commission_split, pc.month_from, pc.month_to
FROM policy_commission pc
WHERE pc.agency_id = %s AND pc.policy_uuid = %s::uuid AND pc.month_from IS NOT NULL
  AND (
      {POLICY_MONTH_EXPR} >= pc.month_from
      AND (
          pc.month_to IS NULL
          OR {POLICY_MONTH_EXPR} <= pc.month_to
      )
  )
"""

POLICY_MONTH_SQL = f"""
SELECT {POLICY_MONTH_EXPR}::int AS policy_month
"""

HISTORICAL_AUTO_RECONCILE_SQL = """
SELECT csr.policy_commission_uuid::text AS policy_commission_uuid
FROM commission_statement_records csr
WHERE csr.agency_id = %s
  AND csr.policy_uuid = %s::uuid
  AND csr.auto_reconcile = true
  AND csr.commission_type = 'OVERRIDE'
  AND csr.policy_commission_uuid IS NOT NULL
  AND csr.is_commission_validated = true
  AND csr.status = 1
  AND (
      CASE WHEN btrim(csr.data->>'gross_commission', E' \\t\\n\\r') ~ '^[+-]?([0-9]+(\\.[0-9]*)?|\\.[0-9]+)$'
          THEN (btrim(csr.data->>'gross_commission', E' \\t\\n\\r'))::NUMERIC
          ELSE NULL
      END
  ) = %s
"""

ALL_SLABS_SQL = """
SELECT month_from, month_to, fee_amount, payout_method, premium_amount, commission_type_text
FROM policy_commission
WHERE agency_id = %s AND policy_uuid = %s::uuid AND month_from IS NOT NULL
ORDER BY month_from, month_to NULLS LAST
"""

FMV_SLABS_SQL = """
SELECT uuid::text AS uuid, month_from, month_to, fee_amount, payout_method, premium_amount,
       commission_type_text, commission_split
FROM policy_commission
WHERE agency_id = %s AND policy_uuid = %s::uuid AND month_from IS NOT NULL
"""

REATTACH_CONTEXT_SQL = """
SELECT
    (
        SELECT MAX(pc.month_to)
        FROM policy_commission pc
        WHERE pc.agency_id = pol.agency_id
          AND pc.policy_uuid = pol.uuid
          AND pc.month_from = 1
          AND pc.commission_type = 1
    ) AS max_month_to,
    prod.uuid::text AS product_uuid,
    pol.premium_amount AS policy_premium_amount,
    prod.attach_latest_commissions AS attach_latest_commissions
FROM policies pol
LEFT JOIN products prod
    ON prod.agency_id = pol.agency_id
   AND prod.uuid = pol.product_id
WHERE pol.agency_id = %s
  AND pol.uuid = %s::uuid
"""

REATTACH_ROWS_SQL = """
SELECT uuid::text AS uuid, commission_type, commission_type_text, month_from, month_to,
       fee_amount, payout_method, premium_amount, commission_split,
       commission_setup_uuid::text AS commission_setup_uuid
FROM policy_commission
WHERE agency_id = %s AND policy_uuid = %s::uuid
"""

FMV_PRIOR_PAYMENTS_SQL = """
SELECT pce.gross_amount, pce.created_at
FROM policy_commission_events pce
INNER JOIN commissions c
    ON c.agency_id = pce.agency_id
   AND c.policy_commission_event_uuid = pce.uuid
WHERE pce.agency_id = %s
  AND pce.policy_uuid = %s::uuid
  AND pce.is_error = false
  AND c.commission_type = %s
ORDER BY pce.created_at DESC
"""

REVIEW_STAGE = 3
REVIEW_FILES_SQL = """
SELECT uuid::text AS uuid, name, agency_id
FROM commission_statements
WHERE stage = %s
  AND agency_id = %s
ORDER BY created_at
"""

SPLIT_BY_UUID_SQL = """
SELECT uuid::text AS uuid, fee_amount, payout_method, commission_split, commission_type_text
FROM policy_commission
WHERE agency_id = %s AND uuid = %s::uuid
"""


def per_member_amount(match, premium):
    fee = money(match.get("fee_amount"))
    if match.get("payout_method") == PAYOUT_PERCENT:
        if premium is None or fee is None:
            return None
        return (premium * fee / Decimal(100)).quantize(MONEY, rounding=ROUND_HALF_UP)
    return fee


def expected_commission_amount(match, premium, members):
    per_member = per_member_amount(match, premium)
    if per_member is None:
        return None
    return (per_member * Decimal(effective_members(members))).quantize(MONEY, rounding=ROUND_HALF_UP)


def slab_window(match):
    start = match.get("month_from")
    end = match.get("month_to")
    end_text = "null" if end is None else str(end)
    return f"month_from {start}, month_to {end_text}"


def describe_existing_commissions(rows, premium, members):
    parts = []
    for row in rows:
        amount = expected_commission_amount(row, premium, members)
        amount_text = format(amount, "f") if amount is not None else "not calculable"
        parts.append(f"{slab_window(row)} amount {amount_text}")
    return "; ".join(parts)


def commission_mismatch_message(target, policy_month, paid_day, effective_day, month_matches, all_slabs, premium, members):
    gross = format(target, "f")
    existing_source = month_matches or all_slabs
    existing = describe_existing_commissions(existing_source, premium, members)
    if policy_month is None:
        return f"Paid-to date or effective date is missing, so the commission month cannot be calculated for gross commission amount {gross}."
    if not month_matches and not all_slabs:
        return f"No commission slab is present for policy month {policy_month} and gross commission amount {gross}."
    if not month_matches:
        return (
            f"No commission slab covers policy month {policy_month} "
            f"(need month_from <= {policy_month} and (month_to is null or month_to >= {policy_month})). "
            f"Effective {effective_day} through paid-to {paid_day}. Gross commission amount {gross}. "
            f"Existing slabs are {existing} only."
        )
    return (
        f"Commission slab for policy month {policy_month} does not match gross commission amount {gross}. "
        f"Effective {effective_day} through paid-to {paid_day}. "
        f"Covering slab is {existing} only."
    )


def is_fmv_product(record):
    value = record.get("commission_processing_type")
    if value is None:
        return False
    try:
        return int(value) == FMV_PROCESSING_TYPE
    except (TypeError, ValueError):
        return False


def remaining_calendar_months(effective_day):
    return 13 - effective_day.month


def fmv_within_tolerance(target, expected):
    return abs(target - expected) <= FMV_TOLERANCE


def fmv_fixed_amount(slab, members):
    fee = money(slab.get("fee_amount"))
    if fee is None:
        return None
    return (fee * Decimal(effective_members(members))).quantize(MONEY, rounding=ROUND_HALF_UP)


def _commission_type_key(slab):
    return str(slab.get("commission_type_text") or "").strip().casefold()


def partition_fmv_slabs(slabs):
    regular = [row for row in slabs if row.get("month_from") is not None and int(row["month_from"]) == 1]
    renewal = [row for row in slabs if row.get("month_from") is not None and int(row["month_from"]) > 12]
    regular.sort(key=lambda row: (row.get("month_to") if row.get("month_to") is not None else 10**9, _commission_type_key(row), str(row.get("uuid") or "")))
    renewal.sort(key=lambda row: (int(row["month_from"]), _commission_type_key(row), str(row.get("uuid") or "")))
    return regular, renewal


def match_fmv_amount(slabs, target, members, effective_day, transaction_type):
    """Same shapes as fileprocessing FmvCommissionMatcher: regular, renewal, NB prorated renewal, remaining initial."""
    regular, renewal = partition_fmv_slabs(slabs)
    remaining = Decimal(remaining_calendar_months(effective_day))
    new_business = str(transaction_type or "").strip().upper() == "NB"

    for slab in regular:
        expected = fmv_fixed_amount(slab, members)
        if expected is not None and fmv_within_tolerance(target, expected):
            return {"rule": "REGULAR", "slab": slab}
    for slab in renewal:
        expected = fmv_fixed_amount(slab, members)
        if expected is not None and fmv_within_tolerance(target, expected):
            return {"rule": "RENEWAL", "slab": slab}
    if new_business:
        for slab in renewal:
            expected = fmv_fixed_amount(slab, members)
            if expected is None:
                continue
            prorated = (expected * remaining).quantize(MONEY, rounding=ROUND_HALF_UP)
            if fmv_within_tolerance(target, prorated):
                return {"rule": "PRORATED_RENEWAL", "slab": slab, "prorated": prorated}
    for regular_slab in regular:
        regular_expected = fmv_fixed_amount(regular_slab, members)
        regular_type = _commission_type_key(regular_slab)
        if regular_expected is None or not regular_type:
            continue
        for renewal_slab in renewal:
            if _commission_type_key(renewal_slab) != regular_type:
                continue
            renewal_expected = fmv_fixed_amount(renewal_slab, members)
            if renewal_expected is None:
                continue
            prorated = (renewal_expected * remaining).quantize(MONEY, rounding=ROUND_HALF_UP)
            remaining_initial = (regular_expected - prorated).quantize(MONEY, rounding=ROUND_HALF_UP)
            if fmv_within_tolerance(target, remaining_initial):
                return {
                    "rule": "REMAINING_INITIAL",
                    "slab": regular_slab,
                    "renewal_slab": renewal_slab,
                    "prorated": prorated,
                }
    return None


def find_prior_prorated_payment(payments, expected):
    qualified = []
    for row in payments:
        amount = money(row.get("gross_amount"))
        if amount is None or not fmv_within_tolerance(amount, expected):
            continue
        created = row.get("created_at")
        recency = -created.timestamp() if hasattr(created, "timestamp") else 0
        qualified.append((abs(amount - expected), recency, amount))
    if not qualified:
        return None
    qualified.sort()
    return qualified[0][2]


def amounts_match(expected, target, payout, tolerance):
    if tolerance and payout == PAYOUT_PERCENT:
        return abs(expected - target) <= PERCENT_TOLERANCE
    return expected == target


def match_amount(matches, target, premium, members):
    member_count = Decimal(effective_members(members))

    def collect(tolerance):
        found = []
        for match in matches:
            per_member = per_member_amount(match, premium)
            if per_member is None:
                continue
            expected = (per_member * member_count).quantize(MONEY, rounding=ROUND_HALF_UP)
            if amounts_match(expected, target, match.get("payout_method"), tolerance):
                found.append(match)
        return found

    exact = collect(False)
    if exact:
        return len(exact), exact[0]
    rounded = collect(True)
    if rounded:
        return len(rounded), rounded[0]
    return 0, None


def agent_matches(record):
    npn = (record.get("agent_npn") or "").strip()
    name = (record.get("agent_name") or "").strip()

    def same_name(column):
        current = record.get(column)
        return AgentNames.matches(name, current)

    checks = [
        (npn and record.get("carrier_agent_appointment_number") == npn, "mapped with carrier_specific_id"),
        (npn and record.get("carrier_agent_npn") == npn, "mapped with npn"),
        (npn and record.get("carrier_agent_user_code") == npn, "mapped with user_code"),
        (same_name("carrier_agent_full_name"), "mapped with agent_name"),
        (npn and record.get("writing_agent_appointment_number") == npn, "mapped with carrier_specific_id"),
        (npn and record.get("writing_agent_npn") == npn, "mapped with npn"),
        (npn and record.get("writing_agent_user_code") == npn, "mapped with user_code"),
        (same_name("writing_agent_full_name"), "mapped with agent_name"),
    ]
    for ok, remark in checks:
        if ok:
            return True, remark
    return False, None


def split_lines(commission_split):
    payload = json_obj(commission_split)
    if not isinstance(payload, list):
        return []
    lines = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        entity = str(item.get("entity") or "").strip().lower()
        if entity not in SPLIT_PARTICIPANTS or not item.get("user_uuid"):
            continue
        lines.append(
            {
                "order": int(item.get("order") or 0),
                "role": item.get("role") or "",
                "entity": entity,
                "user_uuid": str(item.get("user_uuid")),
                "commission_value": money(item.get("commission_value")) or Decimal("0"),
            }
        )
    return lines


def allocate_pool(lines, pool):
    others = Decimal("0")
    built = []
    order_one = None
    for line in lines:
        if line["order"] == 1 and order_one is None:
            order_one = line
            continue
        amount = (pool * line["commission_value"] / Decimal(100)).quantize(MONEY, rounding=ROUND_HALF_UP)
        others += amount
        built.append({**line, "amount": amount})
    if order_one is not None:
        built.append({**order_one, "amount": (pool - others).quantize(MONEY, rounding=ROUND_HALF_UP)})
    return built


def non_prorata_amounts(lines, payout, fee, premium):
    if payout == PAYOUT_FIXED:
        if fee is None or fee <= 0:
            return []
        return allocate_pool(lines, fee.quantize(MONEY, rounding=ROUND_HALF_UP))
    if payout == PAYOUT_PERCENT and premium is not None and fee is not None:
        pool = (premium * fee / Decimal(100)).quantize(MONEY, rounding=ROUND_HALF_UP)
        return allocate_pool(lines, pool)
    return []


def prorata_needed(gross, match, premium):
    fee = money(match.get("fee_amount"))
    if gross is None or fee is None or fee <= 0 or match.get("payout_method") not in (PAYOUT_FIXED, PAYOUT_PERCENT):
        return False
    original = per_member_amount(match, premium)
    return original is not None and gross != original


class Forecast:
    def __init__(self, conn, statement_uuid, auth=None, api_base_url=None):
        self.conn = conn
        self.statement_uuid = statement_uuid
        self.exceptions = []
        self.events = []
        self.commissions = []
        self.details = []
        self.ledgers = {}
        self.ledger_details = []
        self.advance_ledgers = []
        self.records = []
        self.statement = {}
        self.tx = 0
        self.today = date.today()
        self.policies = []
        self.carriers = []
        self.products = []
        self.contacts = []
        self.users = []
        self.agencies = []
        self.existing_details = []
        self.local_policies = []
        self.local_contacts = []
        self.local_commissions = []
        self.reattached_policies = set()
        self.auth = auth if isinstance(auth, dict) else {"token": auth}
        self.api_base_url = (api_base_url or "").rstrip("/")

    def fetch_api(self, url, agency_id):
        try:
            return api_get(url, agency_id, self.auth.get("token"))
        except TokenRejected as rejected:
            logger.warning(
                "API token rejected with status %s. Requesting a new token and retrying this call.",
                rejected.status,
            )
            self.auth["token"] = obtain_token()
            return api_get(url, agency_id, self.auth.get("token"))

    def select(self, sql, params):
        return fetch_all(self.conn, sql, params)

    def next_tx(self, prefix):
        self.tx += 1
        return f"{prefix}-F{self.tx:06d}"

    def add_exception(self, record, code, message):
        record["forecast_status"] = STATUS_FAILED
        record["forecast_error"] = {"code": code, "message": message}
        self.exceptions.append(
            {
                "record_uuid": record["uuid"],
                "policy_number": record.get("policy_number"),
                "row_order": record.get("row_order"),
                "code": code,
                "message": message,
                "match_remarks": record.get("match_remarks"),
            }
        )
        log_row(record, "result=exception code=%s message=%s", code, message)

    def refresh_file_status(self):
        unresolved = [row for row in self.records if row.get("forecast_status") == STATUS_FAILED]
        self.statement["forecast_status"] = FILE_STATUS_WAITING if unresolved else FILE_STATUS_CLEARED
        self.statement["forecast_stage"] = 4 if unresolved else 1

    def reconcile(self):
        rows = self.select(RECORD_SQL, (self.statement_uuid,))
        if not rows:
            raise ForecastError(f"No commission_statement_records found for {self.statement_uuid}")
        rows.sort(key=lambda row: row.get("row_order") or 0)
        logger.info("Loaded %s commission_statement_records for %s", len(rows), self.statement_uuid)
        self.records = rows
        first = rows[0]
        self.statement = {
            "agency_id": first["agency_id"],
            "name": first.get("source_file_name"),
            "batch_id": first.get("source_batch_id"),
            "updated_at": first.get("statement_updated_at"),
            "carrier_uuid": first.get("carrier_uuid"),
        }
        for record in rows:
            record["paid_to_date"] = statement_paid_to_date(record)
            record["gross"] = money(record.get("gross_commission"))
            record["premium_amount"] = money(record.get("premium"))
            record["chargeback_amount"] = money(record.get("chargeback"))
            record["commission_percent_amount"] = money(record.get("commission_percent"))
            if record.get("gross_commission") in (None, "") and record["gross"] is None:
                percent = record["commission_percent_amount"]
                premium = record["premium_amount"]
                if percent is not None and premium is not None:
                    record["gross"] = (premium * percent / Decimal(100)).quantize(MONEY, rounding=ROUND_HALF_UP)
                    record["gross_from_commission_percent"] = True
            record["forecast_status"] = record.get("status")
            record["forecast_policy_uuid"] = record.get("policy_uuid") or record.get("stored_policy_uuid")
            record["forecast_policy_commission_uuid"] = record.get("policy_commission_uuid")
            record["forecast_commission_type"] = record.get("commission_type")
            record["forecast_error"] = None
            record["matched_commission"] = None
            bad_amounts = [
                name
                for name, raw, parsed in (
                    ("commission_percent", record.get("commission_percent"), record["commission_percent_amount"]),
                    ("premium", record.get("premium"), record["premium_amount"]),
                    ("chargeback", record.get("chargeback"), record["chargeback_amount"]),
                )
                if raw not in (None, "") and parsed is None
            ]
            if bad_amounts:
                self.add_exception(
                    record,
                    "INVALID_AMOUNT",
                    "Amount is not a number: "
                    + ", ".join(f"{name}='{record.get(name)}'" for name in bad_amounts),
                )
                continue
            if record.get("status") != STATUS_FAILED:
                log_row(record, "result=skipped status=%s", record.get("status"))
                continue
            self.reconcile_record(record)
        self.refresh_file_status()
        logger.info(
            "Reconcile finished: matched=%s exceptions=%s file_status=%s",
            sum(1 for row in self.records if row.get("forecast_status") == STATUS_MATCHED),
            len(self.exceptions),
            self.statement["forecast_status"],
        )

    def local_month_matches(self, policy_uuid, policy_month):
        if policy_month is None:
            return []
        matches = []
        for row in self.local_commissions:
            if row.get("policy_uuid") != policy_uuid or row.get("month_from") is None:
                continue
            start = int(row["month_from"])
            end = open_month_to(row.get("month_to"))
            if policy_month >= start and (end is None or policy_month <= end):
                matches.append(row)
        return matches

    def local_slabs(self, policy_uuid):
        return [row for row in self.local_commissions if row.get("policy_uuid") == policy_uuid]

    def persisted_commissions(self, sql, params, policy_uuid):
        if policy_uuid and policy_uuid in self.reattached_policies:
            return []
        return self.select(sql, params)

    def find_local_policy(self, record):
        number = str(record.get("policy_number") or "").strip()
        carrier = str(record.get("carrier_uuid") or "")
        for policy in self.local_policies:
            if str(policy.get("number") or "").strip() != number:
                continue
            if carrier and str(policy.get("carrier_id") or "") != carrier:
                continue
            return policy
        return None

    def attach_local_policy(self, record, policy):
        record["policy_exists"] = True
        record["policy_status_id"] = policy.get("status_id", 0)
        record["policy_uuid"] = policy["uuid"]
        record["forecast_policy_uuid"] = policy["uuid"]
        record["policy_effective_date"] = policy.get("effective_date")
        record["pay_to_agent_user_id"] = policy.get("pay_to_agent_user_id")
        record["policy_carrier_id"] = policy.get("carrier_id")
        record["policy_product_uuid"] = policy.get("product_id")
        if policy.get("commission_processing_type") is not None:
            record["commission_processing_type"] = policy.get("commission_processing_type")
        for field in (
            "writing_agent_npn",
            "writing_agent_user_code",
            "writing_agent_appointment_number",
            "writing_agent_full_name",
        ):
            if policy.get(field):
                record[field] = policy[field]

    def resolve_agent_by_npn(self, agency_id, npn, carrier_uuid=None):
        trimmed_npn = (npn or "").strip()
        if not trimmed_npn or agency_id is None:
            return None, "NotFound", []
        rows = self.select(
            AGENT_RESOLVE_SQL,
            (
                agency_id,
                trimmed_npn,
                agency_id,
                trimmed_npn,
                carrier_uuid or "00000000-0000-0000-0000-000000000000",
                agency_id,
                trimmed_npn,
            ),
        )
        valid_rows = [r for r in rows if r.get("agent_uuid")]
        distinct_uuids = list(dict.fromkeys(r["agent_uuid"] for r in valid_rows))
        if not distinct_uuids:
            return None, "NotFound", []
        if len(distinct_uuids) > 1:
            return None, "ConflictingMatches", valid_rows
        return distinct_uuids[0], "Resolved", valid_rows

    def resolve_agent_by_name(self, agency_id, agent_name):
        norm_name = AgentNames.normalize(agent_name)
        if not norm_name or agency_id is None:
            return None, "NotFound", []
        rows = self.select(AGENT_NAME_RESOLVE_SQL, (agency_id, norm_name, norm_name))
        valid_rows = [r for r in rows if r.get("agent_uuid")]
        distinct_uuids = list(dict.fromkeys(r["agent_uuid"] for r in valid_rows))
        if not distinct_uuids:
            return None, "NotFound", []
        if len(distinct_uuids) > 1:
            return None, "AmbiguousFullName", valid_rows
        return distinct_uuids[0], "Resolved", valid_rows

    def check_agent_mismatch(self, record):
        agent_npn = (record.get("agent_npn") or "").strip()
        if not agent_npn and isinstance(record.get("data"), dict):
            agent_npn = str(record["data"].get("agent_npn") or "").strip()
        raw_name = record.get("agent_name")
        if raw_name is None and isinstance(record.get("data"), dict):
            raw_name = record["data"].get("agent_name")
        agent_name = (raw_name or "").strip()
        if not agent_npn or not agent_name:
            return False

        agency_id = record.get("agency_id")
        carrier_uuid = record.get("carrier_uuid") or record.get("policy_carrier_id")

        npn_uuid, npn_status, _ = self.resolve_agent_by_npn(agency_id, agent_npn, carrier_uuid)
        name_uuid, name_status, _ = self.resolve_agent_by_name(agency_id, agent_name)

        if (
            npn_status == "Resolved"
            and name_status == "Resolved"
            and str(npn_uuid).lower() != str(name_uuid).lower()
        ):
            return True
        return False

    def create_local_policy(self, record):
        existing = self.find_local_policy(record)
        if existing:
            self.attach_local_policy(record, existing)
            log_row(record, "result=local_policy_reused policy_uuid=%s", existing["uuid"])
            return True
        details_rows = self.select(
            POLICY_CREATE_DETAILS_SQL,
            (record["agency_id"], record["uuid"], record["file_uuid"]),
        )
        details = details_rows[0] if details_rows else None
        product_name = (details or {}).get("product_name")
        product_uuid = (details or {}).get("product_uuid")
        carrier_uuid = (details or {}).get("carrier_uuid") or record.get("carrier_uuid")
        effective = as_date((details or {}).get("effective_date"))
        agent_npn = (record.get("agent_npn") or "").strip()
        agent_name = (record.get("agent_name") or "").strip()
        if not agent_npn and not agent_name:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                "New policy was not created (reason=agent): agent NPN and agent name are missing.",
            )
            return False

        if self.check_agent_mismatch(record):
            self.add_exception(
                record,
                AGENT_MISMATCH_CODE,
                AGENT_MISMATCH_ID_NAME_MESSAGE,
            )
            return False

        agent_rows = []
        if agent_npn:
            npn_uuid, npn_status, npn_rows = self.resolve_agent_by_npn(
                record["agency_id"],
                agent_npn,
                carrier_uuid,
            )
            if npn_status == "Resolved":
                agent_rows = [npn_rows[0]]
            elif npn_status == "ConflictingMatches":
                self.add_exception(
                    record,
                    "POLICY_MISMATCH",
                    f"New policy was not created (reason=agent): agent NPN '{agent_npn}' was not found.",
                )
                return False

        if not agent_rows and agent_name:
            name_uuid, name_status, name_rows = self.resolve_agent_by_name(
                record["agency_id"],
                agent_name,
            )
            if name_status == "Resolved":
                agent_rows = [name_rows[0]]
            elif name_status == "AmbiguousFullName":
                unique_count = len(set(r["agent_uuid"] for r in name_rows if r.get("agent_uuid")))
                self.add_exception(
                    record,
                    "POLICY_MISMATCH",
                    (
                        f"New policy was not created (reason=agent): agent name '{agent_name}' "
                        f"matched {unique_count} users."
                    ),
                )
                return False
            else:
                self.add_exception(
                    record,
                    "POLICY_MISMATCH",
                    (
                        f"New policy was not created (reason=agent): agent '{agent_name}' "
                        "could not be resolved from NPN, appointment number, user code, or full name."
                    ),
                )
                return False
        if not agent_rows:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                f"New policy was not created (reason=agent): agent NPN '{agent_npn}' was not found.",
            )
            return False
        if not product_uuid:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                (
                    f"New policy was not created (reason=product): product alias "
                    f"'{product_name or ''}' was not found for this carrier."
                ),
            )
            return False
        if effective is None:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                "New policy was not created (reason=effective_date): policy effective date is missing or invalid.",
            )
            return False
        agent_uuid = agent_rows[0]["agent_uuid"]
        resolution = agent_rows[0]["resolution_method"]
        agency_rows = self.select("SELECT uuid::text AS uuid FROM agencies WHERE id = %s", (record["agency_id"],))
        agency_uuid = agency_rows[0]["uuid"] if agency_rows else None
        premium = record.get("premium_amount") or Decimal("0.00")
        match_url = (
            f"{self.api_base_url}/api/v1/commission-nodes/commission-setups/policies/match-commissions"
            f"?product_uuid={product_uuid}&effective_date={effective.isoformat()}&premium={premium}"
        )
        hierarchy_url = (
            f"{self.api_base_url}/api/v1/users/{agent_uuid}/reporting-hierarchy?effective_date={effective.isoformat()}"
        )
        try:
            match_payload = self.fetch_api(match_url, record["agency_id"])
            hierarchy_payload = self.fetch_api(hierarchy_url, record["agency_id"])
        except ForecastError as exc:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                f"New policy was not created (reason=commission): commission lookup failed: {exc}",
            )
            return False
        if isinstance(hierarchy_payload, dict) and (hierarchy_payload.get("data") or [None])[0] and isinstance(
            (hierarchy_payload.get("data") or [None])[0], dict
        ) and (hierarchy_payload.get("data") or [{}])[0].get("message"):
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                (
                    "New policy was not created (reason=commission): reporting hierarchy returned: "
                    f"{(hierarchy_payload.get('data') or [{}])[0].get('message')}"
                ),
            )
            return False
        hierarchy_users = []
        for user in (hierarchy_payload or {}).get("data") or []:
            if isinstance(user, dict) and user.get("uuid"):
                hierarchy_users.append(user)
        if not hierarchy_users:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                "New policy was not created (reason=commission): reporting hierarchy returned no users.",
            )
            return False
        hierarchy_chain = build_hierarchy_chain(hierarchy_users, agent_uuid)
        if not hierarchy_chain:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                (
                    "New policy was not created (reason=commission): reporting hierarchy data missing "
                    f"pay-to agent {agent_uuid}."
                ),
            )
            return False
        affiliated = hierarchy_chain[0].get("affiliatedAgencyUuid") or hierarchy_chain[0].get(
            "affiliated_agency_uuid"
        )
        hierarchy_agency_uuid = str(affiliated) if affiliated else agency_uuid
        if not hierarchy_agency_uuid:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                "New policy was not created (reason=commission): tenant agency UUID is missing.",
            )
            return False
        try:
            agency_hierarchy_payload = self.fetch_api(
                f"{self.api_base_url}/api/v1/agencies/hierarchy",
                record["agency_id"],
            )
        except ForecastError as exc:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                f"New policy was not created (reason=commission): agency hierarchy lookup failed: {exc}",
            )
            return False
        agency_levels = agency_hierarchy_levels(agency_hierarchy_payload, hierarchy_agency_uuid)
        if not agency_levels:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                (
                    "New policy was not created (reason=commission): agency hierarchy returned no levels "
                    f"for {hierarchy_agency_uuid}."
                ),
            )
            return False
        level_maps = build_level_order_maps(agency_levels)
        policy_uuid = new_id()
        slabs = slabs_from_match_response(match_payload, policy_uuid)
        if not slabs:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                (
                    "New policy was not created (reason=commission): match-commission returned no "
                    f"published commission for product {product_uuid} on {effective.isoformat()}."
                ),
            )
            return False
        built_slabs = []
        try:
            for slab in slabs:
                period = slab.pop("period")
                slab["commission_split"] = build_commission_splits_for_period(
                    period,
                    agent_uuid,
                    agency_uuid,
                    hierarchy_users,
                    level_maps,
                )
                built_slabs.append(slab)
        except ForecastError as exc:
            self.add_exception(
                record,
                "POLICY_MISMATCH",
                f"New policy was not created (reason=commission): {exc}",
            )
            return False
        self.local_commissions.extend(built_slabs)
        effective_dt = datetime.combine(effective, datetime.min.time())
        local_policy = {
            "uuid": policy_uuid,
            "agency_id": record["agency_id"],
            "number": record.get("policy_number"),
            "effective_date": effective_dt,
            "carrier_id": carrier_uuid,
            "product_id": product_uuid,
            "status_id": 0,
            "pay_to_agent_user_id": agent_uuid,
        }
        self.local_policies.append(local_policy)
        self.local_contacts.append(
            {
                "agency_id": record["agency_id"],
                "relationship_id": policy_uuid,
                "first_name": (details or {}).get("holder_first_name"),
                "last_name": (details or {}).get("holder_last_name"),
            }
        )
        record["policy_exists"] = True
        record["policy_status_id"] = 0
        record["policy_uuid"] = policy_uuid
        record["forecast_policy_uuid"] = policy_uuid
        record["policy_effective_date"] = effective_dt
        record["pay_to_agent_user_id"] = agent_uuid
        record["policy_carrier_id"] = carrier_uuid
        record["policy_product_uuid"] = product_uuid
        if product_uuid:
            type_rows = self.select(
                "SELECT commission_processing_type FROM products WHERE agency_id = %s AND uuid = %s::uuid",
                (record["agency_id"], product_uuid),
            )
            if type_rows:
                record["commission_processing_type"] = type_rows[0].get("commission_processing_type")
        if resolution == "npn":
            record["writing_agent_npn"] = agent_npn
        elif resolution == "user_code":
            record["writing_agent_user_code"] = agent_npn
        elif resolution == "carrier_appointment":
            record["writing_agent_appointment_number"] = agent_npn
        else:
            record["writing_agent_full_name"] = agent_name
        local_policy["commission_processing_type"] = record.get("commission_processing_type")
        for field in (
            "writing_agent_npn",
            "writing_agent_user_code",
            "writing_agent_appointment_number",
            "writing_agent_full_name",
        ):
            if record.get(field):
                local_policy[field] = record[field]
        log_row(
            record,
            "result=local_policy_created policy_uuid=%s product=%s slabs=%s agent_resolved=%s",
            policy_uuid,
            product_uuid,
            len(slabs),
            resolution,
        )
        return True

    def reattach_context(self, record):
        policy_uuid = record.get("forecast_policy_uuid")
        local_ids = {policy.get("uuid") for policy in self.local_policies}
        if policy_uuid not in local_ids:
            rows = self.select(REATTACH_CONTEXT_SQL, (record["agency_id"], policy_uuid))
            if rows:
                return rows[0]
        product_uuid = record.get("policy_product_uuid")
        if not product_uuid:
            return None
        flag_rows = self.select(
            "SELECT attach_latest_commissions FROM products WHERE agency_id = %s AND uuid = %s::uuid",
            (record["agency_id"], product_uuid),
        )
        if not flag_rows:
            return None
        ends = []
        for row in self.local_slabs(policy_uuid):
            if row.get("month_from") is None or int(row["month_from"]) != REATTACH_MONTH_FROM:
                continue
            if slab_type_code(row) != COMMISSION_TYPE_REGULAR:
                continue
            if row.get("month_to") is not None:
                ends.append(int(row["month_to"]))
        return {
            "max_month_to": max(ends) if ends else None,
            "product_uuid": product_uuid,
            "policy_premium_amount": None,
            "attach_latest_commissions": flag_rows[0].get("attach_latest_commissions"),
        }

    def commission_rows_for_reattach(self, record):
        policy_uuid = record.get("forecast_policy_uuid")
        local = self.local_slabs(policy_uuid)
        if local or policy_uuid in self.reattached_policies:
            return local
        rows = self.select(REATTACH_ROWS_SQL, (record["agency_id"], policy_uuid))
        for row in rows:
            row["commission_split"] = json_obj(row.get("commission_split")) or []
            row["month_to"] = open_month_to(row.get("month_to"))
        return rows

    def fetch_match_slabs(self, record, product_uuid, agent_uuid, hierarchy_day, match_day, premium):
        """Latest product slabs. Hierarchy uses the policy effective date; match-commission uses paid-to."""
        match_url = (
            f"{self.api_base_url}/api/v1/commission-nodes/commission-setups/policies/match-commissions"
            f"?product_uuid={product_uuid}&effective_date={match_day.isoformat()}&premium={premium}"
        )
        hierarchy_url = (
            f"{self.api_base_url}/api/v1/users/{agent_uuid}/reporting-hierarchy?effective_date={hierarchy_day.isoformat()}"
        )
        match_payload = self.fetch_api(match_url, record["agency_id"])
        hierarchy_payload = self.fetch_api(hierarchy_url, record["agency_id"])
        hierarchy_data = (hierarchy_payload or {}).get("data") or []
        if hierarchy_data and isinstance(hierarchy_data[0], dict) and hierarchy_data[0].get("message"):
            raise ForecastError(hierarchy_data[0].get("message"))
        hierarchy_users = [user for user in hierarchy_data if isinstance(user, dict) and user.get("uuid")]
        if not hierarchy_users:
            raise ForecastError("reporting hierarchy returned no users")
        hierarchy_chain = build_hierarchy_chain(hierarchy_users, str(agent_uuid))
        if not hierarchy_chain:
            raise ForecastError(f"reporting hierarchy missing pay-to agent {agent_uuid}")
        agency_rows = self.select("SELECT uuid::text AS uuid FROM agencies WHERE id = %s", (record["agency_id"],))
        agency_uuid = agency_rows[0]["uuid"] if agency_rows else None
        affiliated = hierarchy_chain[0].get("affiliatedAgencyUuid") or hierarchy_chain[0].get("affiliated_agency_uuid")
        hierarchy_agency_uuid = str(affiliated) if affiliated else agency_uuid
        if not hierarchy_agency_uuid:
            raise ForecastError("tenant agency UUID is missing")
        agency_hierarchy_payload = self.fetch_api(
            f"{self.api_base_url}/api/v1/agencies/hierarchy",
            record["agency_id"],
        )
        agency_levels = agency_hierarchy_levels(agency_hierarchy_payload, hierarchy_agency_uuid)
        if not agency_levels:
            raise ForecastError(f"agency hierarchy returned no levels for {hierarchy_agency_uuid}")
        level_maps = build_level_order_maps(agency_levels)
        slabs = slabs_from_match_response(match_payload, record.get("forecast_policy_uuid"))
        built = []
        for slab in slabs:
            period = slab.pop("period")
            slab["commission_split"] = build_commission_splits_for_period(
                period,
                str(agent_uuid),
                agency_uuid,
                hierarchy_users,
                level_maps,
            )
            built.append(slab)
        return built

    def maybe_reattach_commissions(self, record):
        """Replace stored slabs with the latest product setup when the product flag is on and the policy month is past the current slab end."""
        policy_uuid = record.get("forecast_policy_uuid")
        if not policy_uuid or policy_uuid in self.reattached_policies:
            return
        if record.get("policy_commission_uuid") or record.get("chargeback_uuid"):
            return
        pay_to = record.get("pay_to_agent_user_id")
        effective = calendar_day(record.get("policy_effective_date"))
        paid = calendar_day(record.get("paid_to_date"))
        if not pay_to or effective is None or paid is None:
            return
        month_index = policy_month_index(paid, effective)
        if month_index is None:
            return
        context = self.reattach_context(record)
        if not context:
            return
        max_month = context.get("max_month_to")
        if max_month is None:
            max_month = REATTACH_DEFAULT_MAX_MONTH_TO
        if month_index <= int(max_month):
            return
        if not as_bool(context.get("attach_latest_commissions")):
            return
        product_uuid = context.get("product_uuid") or record.get("policy_product_uuid")
        if not product_uuid:
            log_row(record, "reattach_skipped reason=product")
            return
        premium = record.get("premium_amount")
        if premium is None:
            premium = money(context.get("policy_premium_amount")) or Decimal("0.00")
        try:
            incoming = self.fetch_match_slabs(record, product_uuid, pay_to, effective, paid, premium)
        except ForecastError as exc:
            log_row(record, "reattach_skipped reason=%s", exc)
            return
        for slab in incoming:
            slab["policy_uuid"] = policy_uuid
        existing = self.commission_rows_for_reattach(record)
        if not existing:
            self.local_commissions = [row for row in self.local_commissions if row.get("policy_uuid") != policy_uuid]
            self.local_commissions.extend(incoming)
            self.reattached_policies.add(policy_uuid)
            log_row(record, "result=commissions_reattached slabs=%s reason=created", len(incoming))
            return
        actions = plan_reattach_actions(existing, incoming)
        if not actions:
            log_row(record, "reattach_skipped reason=already_latest")
            return
        rewritten = apply_reattach_actions(existing, actions)
        for slab in rewritten:
            slab["policy_uuid"] = policy_uuid
        self.local_commissions = [row for row in self.local_commissions if row.get("policy_uuid") != policy_uuid]
        self.local_commissions.extend(rewritten)
        self.reattached_policies.add(policy_uuid)
        log_row(record, "result=commissions_reattached slabs=%s", len(rewritten))

    def reconcile_record(self, record):
        if not record.get("policy_number"):
            self.add_exception(record, "POLICY_MISMATCH", "Missing policy number")
            return
        if not record.get("policy_exists"):
            if not self.create_local_policy(record):
                return
        status_id = record.get("policy_status_id")
        if status_id is None:
            self.add_exception(record, "POLICY_MISMATCH", "Policy found but not active (status_id=null, unknown)")
            return
        if int(status_id) != POLICY_ACTIVE:
            label = POLICY_STATUS_LABELS.get(int(status_id), str(status_id))
            self.add_exception(record, "POLICY_MISMATCH", f"Policy found but not active ({label})")
            return
        if self.check_agent_mismatch(record):
            self.add_exception(record, AGENT_MISMATCH_CODE, AGENT_MISMATCH_ID_NAME_MESSAGE)
            return
        matched, remark = agent_matches(record)
        if not matched:
            self.add_exception(record, "AGENT_MISMATCH", "Agent not found")
            return
        record["match_remarks"] = json.dumps({"agent": remark})
        record["forecast_policy_uuid"] = record.get("policy_uuid")
        self.maybe_reattach_commissions(record)
        if self.cancellation_gate(record):
            self.add_exception(record, "POLICY_CANCELLATION", "Policy cancellation required")
            return
        gross = record["gross"]
        if self.gross_has_value(gross) and gross < 0:
            if as_bool(record.get("is_chargeback_validated")) or self.try_chargeback_auto_resolve(record):
                record["forecast_status"] = STATUS_MATCHED
                log_row(record, "result=matched reason=chargeback")
                return
            self.add_exception(record, self.chargeback_failure_code(record), "Gross commission is negative")
            return
        if as_bool(record.get("is_commission_validated")) or (
            as_bool(record.get("is_advance_validated")) and not as_bool(record.get("is_commission_validated"))
        ):
            if self.settle_chargeback(record):
                record["forecast_status"] = STATUS_MATCHED
                log_row(record, "result=matched reason=already_validated")
            return
        if not self.gross_has_value(gross):
            if self.chargeback_has_value(record.get("chargeback_amount")):
                if self.settle_chargeback(record):
                    record["forecast_status"] = STATUS_MATCHED
                    log_row(record, "result=matched reason=chargeback")
                return
            if gross is None:
                self.add_exception(
                    record,
                    "MISSING_GROSS",
                    "Gross commission is missing, and it cannot be calculated because premium or commission_percent is missing.",
                )
                return
        if record.get("gross_from_commission_percent"):
            log_row(
                record,
                "gross_calculated_from_commission_percent percent=%s",
                record.get("commission_percent"),
            )
        self.match_commission(record, gross.copy_abs() if gross is not None and gross < 0 else gross)
        if record.get("forecast_status") == STATUS_MATCHED:
            self.settle_chargeback(record)

    def gross_has_value(self, gross):
        return gross is not None and gross != 0

    def chargeback_has_value(self, chargeback):
        return chargeback is not None and chargeback != 0

    def chargeback_match_amount(self, record):
        chargeback = record.get("chargeback_amount")
        if self.chargeback_has_value(chargeback):
            return abs(chargeback)
        gross = record.get("gross")
        if gross is not None and gross < 0:
            return abs(gross)
        return None

    def has_active_arf(self, record):
        if not record.get("has_policy_advance"):
            return False
        try:
            status = int(record.get("policy_advance_status"))
        except (TypeError, ValueError):
            return False
        if status != 1:
            return False
        advance = money(record.get("policy_advance_amount"))
        recovered = money(record.get("policy_advance_recovered_amount"))
        if advance is None or recovered is None:
            return False
        return advance > recovered

    def chargeback_failure_code(self, record):
        if as_bool(record.get("is_policy_cancellation_validated")):
            return "CHARGEBACK_RECOVERY"
        gross = record.get("gross")
        if (
            as_bool(record.get("is_advance_validated"))
            and gross is not None
            and gross > 0
            and self.chargeback_has_value(record.get("chargeback_amount"))
        ):
            return "CHARGEBACK_RECOVERY"
        initiated = str(record.get("policy_advance_initiated_by") or "").strip().lower()
        if self.has_active_arf(record) and initiated == "carrier":
            return "POLICY_CANCELLATION"
        return "CHARGEBACK_RECOVERY"

    def period_matches(self, record):
        paid = record.get("paid_to_date") or record.get("policy_effective_date")
        effective = record.get("policy_effective_date")
        if paid is None or effective is None or not record.get("forecast_policy_uuid"):
            return []
        paid_day = paid.date() if isinstance(paid, datetime) else paid
        effective_day = effective.date() if isinstance(effective, datetime) else effective
        matches = self.persisted_commissions(
            MATCH_SQL,
            (
                record["agency_id"],
                record["forecast_policy_uuid"],
                paid_day,
                effective_day,
                paid_day,
                effective_day,
                paid_day,
                effective_day,
                paid_day,
                effective_day,
            ),
            record.get("forecast_policy_uuid"),
        )
        month_row = self.select(POLICY_MONTH_SQL, (paid_day, effective_day, paid_day, effective_day))
        policy_month = month_row[0]["policy_month"] if month_row else None
        return self.local_month_matches(record.get("forecast_policy_uuid"), policy_month) + matches

    def try_chargeback_auto_resolve(self, record):
        if as_bool(record.get("is_chargeback_validated")) or record.get("chargeback_uuid"):
            return False
        if as_bool(record.get("is_policy_cancellation_validated")):
            return False
        if not self.has_active_arf(record):
            return False
        if not record.get("forecast_policy_uuid"):
            return False
        amount = self.chargeback_match_amount(record)
        if amount is None:
            return False
        paid = record.get("paid_to_date") or record.get("policy_effective_date")
        effective = record.get("policy_effective_date")
        if paid is None or effective is None:
            return False
        matches = self.period_matches(record)
        count, chosen = match_amount(matches, amount, record.get("premium_amount"), record.get("total_members"))
        if count != 1 or chosen is None:
            return False
        record["chargeback_uuid"] = chosen["uuid"]
        record["chargeback_entity_type"] = "split"
        record["is_chargeback_validated"] = True
        record["matched_commission"] = chosen
        existing_type = str(record.get("forecast_commission_type") or record.get("commission_type") or "").strip()
        if existing_type and "CHARGEBACK" not in existing_type.upper():
            record["forecast_commission_type"] = f"{existing_type}, CHARGEBACK"
        else:
            record["forecast_commission_type"] = existing_type or TYPE_CHARGEBACK
        log_row(
            record,
            "result=chargeback_resolved commission_uuid=%s amount=%s",
            chosen["uuid"],
            amount,
        )
        return True

    def settle_chargeback(self, record):
        if not self.chargeback_has_value(record.get("chargeback_amount")):
            return True
        if as_bool(record.get("is_chargeback_validated")) or self.try_chargeback_auto_resolve(record):
            return True
        self.add_exception(
            record,
            self.chargeback_failure_code(record),
            f"Chargeback amount {format(abs(record['chargeback_amount']), 'f')} was not matched to one commission slab.",
        )
        return False

    def try_historical_auto_reconcile(self, record, matches, target):
        if not matches or target is None or not record.get("forecast_policy_uuid"):
            return False
        rows = self.select(
            HISTORICAL_AUTO_RECONCILE_SQL,
            (record["agency_id"], record["forecast_policy_uuid"], target),
        )
        historical = {row["policy_commission_uuid"] for row in rows if row.get("policy_commission_uuid")}
        candidates = [
            match
            for match in matches
            if match.get("uuid") in historical
            and str(match.get("commission_type_text") or "").strip().upper() == "OVERRIDE"
        ]
        if len(candidates) == 1:
            chosen = candidates[0]
            record["forecast_policy_commission_uuid"] = chosen["uuid"]
            record["forecast_commission_type"] = chosen.get("commission_type_text") or TYPE_COMMISSION
            record["matched_commission"] = chosen
            record["is_commission_validated"] = True
            record["forecast_status"] = STATUS_MATCHED
            log_row(
                record,
                "result=matched commission_uuid=%s type=%s reason=historical_override",
                chosen["uuid"],
                record["forecast_commission_type"],
            )
            return True
        if len(candidates) > 1:
            existing = describe_existing_commissions(candidates, record.get("premium_amount"), record.get("total_members"))
            self.add_exception(
                record,
                "MULTIPLE_COMMISSIONS",
                f"More than one historical override matches gross commission amount {format(target, 'f')}. Existing amounts are {existing}.",
            )
            return True
        return False

    def accept_commission_match(self, record, chosen, rule=None):
        record["forecast_policy_commission_uuid"] = chosen["uuid"]
        record["forecast_commission_type"] = chosen.get("commission_type_text") or TYPE_COMMISSION
        record["matched_commission"] = chosen
        record["is_commission_validated"] = True
        record["forecast_status"] = STATUS_MATCHED
        if rule:
            record["fmv_match_rule"] = rule
            log_row(
                record,
                "result=matched commission_uuid=%s type=%s fmv=%s",
                chosen["uuid"],
                record["forecast_commission_type"],
                rule,
            )
            return
        log_row(
            record,
            "result=matched commission_uuid=%s type=%s",
            chosen["uuid"],
            record["forecast_commission_type"],
        )

    def fmv_slabs_for(self, record):
        policy_uuid = record.get("forecast_policy_uuid")
        slabs = list(self.local_slabs(policy_uuid))
        slabs.extend(self.persisted_commissions(FMV_SLABS_SQL, (record["agency_id"], policy_uuid), policy_uuid))
        return slabs

    def prior_fmv_payments(self, record, commission_type, expected):
        payments = self.select(
            FMV_PRIOR_PAYMENTS_SQL,
            (record["agency_id"], record.get("forecast_policy_uuid"), commission_type),
        )
        for earlier in self.records:
            if earlier is record:
                break
            if earlier.get("fmv_match_rule") != "PRORATED_RENEWAL":
                continue
            if earlier.get("forecast_policy_uuid") != record.get("forecast_policy_uuid"):
                continue
            earlier_type = str(earlier.get("forecast_commission_type") or "").strip().upper()
            if earlier_type != commission_type:
                continue
            payments.append({"gross_amount": earlier.get("gross"), "created_at": None})
        return find_prior_prorated_payment(payments, expected)

    def match_fmv_commission(self, record, target, effective_day):
        slabs = self.fmv_slabs_for(record)
        result = match_fmv_amount(
            slabs,
            target,
            record.get("total_members"),
            effective_day,
            record.get("transaction_type"),
        )
        if result is None:
            self.add_exception(
                record,
                "COMMISSION_MISMATCH",
                (
                    f"FMV commission amount {format(target, 'f')} does not match Regular, Renewal, "
                    f"prorated Renewal, or Regular minus prorated Renewal."
                ),
            )
            return
        if result["rule"] != "REMAINING_INITIAL":
            self.accept_commission_match(record, result["slab"], result["rule"])
            return
        commission_type = str(result["slab"].get("commission_type_text") or result["renewal_slab"].get("commission_type_text") or "").strip().upper()
        if not commission_type:
            self.add_exception(
                record,
                "TRUE_UP_MISMATCH",
                "FMV remaining initial commission matched but no offsetting renewal true-up transaction was found",
            )
            return
        paid = self.prior_fmv_payments(record, commission_type, result["prorated"])
        if paid is None:
            self.add_exception(
                record,
                "TRUE_UP_MISMATCH",
                "FMV remaining initial commission matched but no offsetting renewal true-up transaction was found",
            )
            return
        self.accept_commission_match(record, result["slab"], result["rule"])
        record["chargeback_amount"] = paid
        record["chargeback_uuid"] = result["renewal_slab"].get("uuid")
        record["chargeback_entity_type"] = "split"
        record["is_chargeback_validated"] = True
        record["true_up_adjustment"] = paid
        record["is_true_up_validated"] = True
        log_row(record, "fmv_true_up amount=%s renewal_slab=%s", paid, record["chargeback_uuid"])

    def cancellation_gate(self, record):
        chargeback = record["chargeback_amount"]
        gross = record["gross"]
        is_chargeback = (chargeback is not None and chargeback != 0) or (gross is not None and gross < 0)
        if not is_chargeback or as_bool(record.get("is_policy_cancellation_validated")) or not record.get("has_policy_advance"):
            return False
        return str(record.get("policy_advance_initiated_by") or "").strip().lower() == "agency"

    def explain_commission_mismatch(self, record, target):
        paid = record.get("paid_to_date") or record.get("policy_effective_date")
        effective = record.get("policy_effective_date")
        premium = record.get("premium_amount")
        members = record.get("total_members")
        if target is None:
            self.add_exception(record, "COMMISSION_MISMATCH", "Gross commission amount is missing, so no commission record can be matched.")
            return
        if paid is None or effective is None or not record.get("forecast_policy_uuid"):
            self.add_exception(
                record,
                "COMMISSION_MISMATCH",
                commission_mismatch_message(target, None, None, None, [], [], premium, members),
            )
            return
        paid_day = paid.date() if isinstance(paid, datetime) else paid
        effective_day = effective.date() if isinstance(effective, datetime) else effective
        month_row = self.select(POLICY_MONTH_SQL, (paid_day, effective_day, paid_day, effective_day))
        policy_month = month_row[0]["policy_month"] if month_row else None
        month_matches = self.persisted_commissions(
            MATCH_SQL,
            (
                record["agency_id"],
                record["forecast_policy_uuid"],
                paid_day,
                effective_day,
                paid_day,
                effective_day,
                paid_day,
                effective_day,
                paid_day,
                effective_day,
            ),
            record.get("forecast_policy_uuid"),
        )
        month_matches = self.local_month_matches(record.get("forecast_policy_uuid"), policy_month) + month_matches
        all_slabs = []
        if not month_matches:
            all_slabs = self.local_slabs(record.get("forecast_policy_uuid")) + self.persisted_commissions(
                ALL_SLABS_SQL,
                (record["agency_id"], record["forecast_policy_uuid"]),
                record.get("forecast_policy_uuid"),
            )
        self.add_exception(
            record,
            "COMMISSION_MISMATCH",
            commission_mismatch_message(
                target,
                policy_month,
                paid_day,
                effective_day,
                month_matches,
                all_slabs,
                premium,
                members,
            ),
        )
        return month_matches

    def match_commission(self, record, target):
        paid = record.get("paid_to_date") or record.get("policy_effective_date")
        effective = record.get("policy_effective_date")
        if paid is None or effective is None or not record.get("forecast_policy_uuid"):
            self.explain_commission_mismatch(record, target)
            return
        paid_day = paid.date() if isinstance(paid, datetime) else paid
        effective_day = effective.date() if isinstance(effective, datetime) else effective
        if is_fmv_product(record):
            self.match_fmv_commission(record, target, effective_day)
            return
        matches = self.persisted_commissions(
            MATCH_SQL,
            (
                record["agency_id"],
                record["forecast_policy_uuid"],
                paid_day,
                effective_day,
                paid_day,
                effective_day,
                paid_day,
                effective_day,
                paid_day,
                effective_day,
            ),
            record.get("forecast_policy_uuid"),
        )
        month_row = self.select(POLICY_MONTH_SQL, (paid_day, effective_day, paid_day, effective_day))
        policy_month = month_row[0]["policy_month"] if month_row else None
        matches = self.local_month_matches(record.get("forecast_policy_uuid"), policy_month) + matches
        count, chosen = match_amount(matches, target, record["premium_amount"], record.get("total_members"))
        if count == 0:
            if self.try_historical_auto_reconcile(record, matches, target):
                return
            self.explain_commission_mismatch(record, target)
            return
        if count > 1:
            existing = describe_existing_commissions(matches, record["premium_amount"], record.get("total_members"))
            self.add_exception(
                record,
                "MULTIPLE_COMMISSIONS",
                f"More than one commission record matches gross commission amount {format(target, 'f')}. Existing amounts are {existing}.",
            )
            return
        record["forecast_policy_commission_uuid"] = chosen["uuid"]
        record["forecast_commission_type"] = chosen.get("commission_type_text") or TYPE_COMMISSION
        record["matched_commission"] = chosen
        record["is_commission_validated"] = True
        record["forecast_status"] = STATUS_MATCHED
        log_row(
            record,
            "result=matched commission_uuid=%s type=%s",
            chosen["uuid"],
            record["forecast_commission_type"],
        )

    def process_commissions(self):
        for record in self.records:
            if record.get("forecast_status") == STATUS_MATCHED and record.get("status") != STATUS_MATCHED:
                self.process_record(record)
        self.refresh_file_status()
        logger.info(
            "Commission forecast finished: events=%s commissions=%s details=%s ledger_lines=%s",
            len(self.events),
            len(self.commissions),
            len(self.details),
            sum(1 for row in self.ledger_details if not row.get("id")),
        )

    def process_record(self, record):
        policy_uuid = record.get("forecast_policy_uuid")
        if not policy_uuid:
            return
        event_uuid = new_id()
        self.events.append(
            {
                "uuid": event_uuid,
                "agency_id": record["agency_id"],
                "statement_uuid": record["uuid"],
                "policy_uuid": policy_uuid,
                "policy_commission_uuid": record.get("forecast_policy_commission_uuid"),
                "gross_amount": record["gross"],
                "status": "PROCESSED",
            }
        )
        commission_validated = as_bool(record.get("is_commission_validated"))
        chargeback_validated = as_bool(record.get("is_chargeback_validated"))
        gross = record["gross"]
        try:
            if not commission_validated and not chargeback_validated:
                if gross is not None:
                    self.add_commission(record, event_uuid, record.get("forecast_commission_type") or TYPE_COMMISSION, gross)
                return
            advance_amount = money(record.get("policy_advance_amount"))
            if (
                as_bool(record.get("policy_advance_process_advance"))
                and as_bool(record.get("is_advance_validated"))
                and advance_amount
                and advance_amount > 0
            ):
                self.apply_advance(record, event_uuid, advance_amount)
                return
            match = record.get("matched_commission")
            if match is None and record.get("forecast_policy_commission_uuid"):
                found = self.select(SPLIT_BY_UUID_SQL, (record["agency_id"], record["forecast_policy_commission_uuid"]))
                match = found[0] if found else None
            if commission_validated and gross is not None and gross > 0 and match is not None:
                self.apply_commission(record, event_uuid, match, gross, TYPE_COMMISSION)
            true_up = money(record.get("true_up_adjustment"))
            if as_bool(record.get("is_true_up_validated")) and true_up and true_up > 0 and match is not None:
                self.apply_commission(record, event_uuid, match, true_up, TYPE_TRUE_UP)
            if chargeback_validated and (
                self.chargeback_has_value(record.get("chargeback_amount")) or (gross is not None and gross < 0)
            ):
                chargeback = record.get("chargeback_amount")
                amount = abs(chargeback) if self.chargeback_has_value(chargeback) else abs(gross or 0)
                self.apply_chargeback(record, event_uuid, amount)
        finally:
            log_row(
                record,
                "result=commission_processed event=%s commission_validated=%s chargeback_validated=%s",
                event_uuid,
                commission_validated,
                chargeback_validated,
            )

    def apply_advance(self, record, event_uuid, amount):
        entity_type = str(record.get("policy_advance_entity_type") or ENTITY_AGENT).strip().lower()
        entity_uuid = record.get("policy_advance_entity_uuid") or record.get("pay_to_agent_user_id")
        if not entity_uuid:
            self.add_exception(record, "ADVANCE", "Advance is marked for processing but the payee is missing")
            return
        if entity_type == ENTITY_AGENCY:
            commission_uuid = self.add_commission(record, event_uuid, TYPE_AGENCY_CREDIT, amount)
            self.persist_entry(
                record,
                commission_uuid,
                TYPE_AGENCY_CREDIT,
                {"entity": ENTITY_AGENCY, "user_uuid": entity_uuid, "role": "Agency", "amount": amount},
            )
            return
        commission_uuid = self.add_commission(record, event_uuid, TYPE_ARF, amount)
        self.persist_entry(
            record,
            commission_uuid,
            TYPE_ARF,
            {"entity": ENTITY_AGENT, "user_uuid": entity_uuid, "role": "Agent", "amount": amount},
            ledger=False,
        )
        self.advance_ledgers.append(
            {
                "uuid": new_id(),
                "agency_id": record["agency_id"],
                "policy_uuid": record.get("forecast_policy_uuid"),
                "entity_uuid": entity_uuid,
                "amount": amount,
                "transaction_id": self.next_tx("ADV"),
            }
        )

    def apply_commission(self, record, event_uuid, match, gross, commission_type):
        lines = split_lines(match.get("commission_split"))
        if commission_type == TYPE_TRUE_UP:
            allocated, header = allocate_pool(lines, gross), gross
        elif prorata_needed(gross, match, record["premium_amount"]):
            if not any(line["order"] == 1 for line in lines):
                self.add_exception(record, "COMMISSION_MISMATCH", "Commission amount or period not matched")
                return
            allocated, header = allocate_pool(lines, gross), gross
        else:
            allocated = non_prorata_amounts(
                lines,
                match.get("payout_method"),
                money(match.get("fee_amount")),
                record["premium_amount"],
            )
            header = per_member_amount(match, record["premium_amount"]) or gross
        commission_uuid = self.add_commission(record, event_uuid, commission_type, header)
        for line in allocated:
            self.persist_entry(record, commission_uuid, commission_type, line)

    def apply_chargeback(self, record, event_uuid, amount):
        entity_type = str(record.get("chargeback_entity_type") or "").strip().lower()
        negative = amount.copy_abs() * Decimal("-1")
        if entity_type == "split" and record.get("chargeback_uuid"):
            found = self.select(SPLIT_BY_UUID_SQL, (record["agency_id"], record["chargeback_uuid"]))
            if not found:
                return
            lines = split_lines(found[0].get("commission_split"))
            if not any(line["order"] == 1 for line in lines):
                return
            allocated = allocate_pool(lines, negative)
            commission_uuid = self.add_commission(record, event_uuid, TYPE_CHARGEBACK, negative)
            for line in allocated:
                self.persist_entry(record, commission_uuid, TYPE_CHARGEBACK, line)
            return
        if not record.get("chargeback_uuid"):
            return
        mapped = entity_type if entity_type in SPLIT_PARTICIPANTS else ENTITY_AGENT
        commission_uuid = self.add_commission(record, event_uuid, TYPE_CHARGEBACK, negative)
        self.persist_entry(
            record,
            commission_uuid,
            TYPE_CHARGEBACK,
            {"entity": mapped, "user_uuid": record["chargeback_uuid"], "role": mapped, "amount": negative},
        )

    def add_commission(self, record, event_uuid, commission_type, amount):
        commission_uuid = new_id()
        self.commissions.append(
            {
                "uuid": commission_uuid,
                "agency_id": record["agency_id"],
                "policy_commission_event_uuid": event_uuid,
                "commission_type": commission_type,
                "amount": money(amount),
                "transaction_id": self.next_tx("POL"),
                "statement_uuid": record["uuid"],
                "policy_uuid": record.get("forecast_policy_uuid"),
                "policy_number": record.get("policy_number"),
            }
        )
        return commission_uuid

    def persist_entry(self, record, commission_uuid, commission_type, line, ledger=True):
        amount = money(line["amount"]) or Decimal("0.00")
        entity = str(line["entity"]).lower()
        detail_uuid = new_id()
        zero = amount == 0
        self.details.append(
            {
                "uuid": detail_uuid,
                "agency_id": record["agency_id"],
                "commission_uuid": commission_uuid,
                "entity_uuid": line["user_uuid"],
                "entity_role": line.get("role") or entity,
                "entity_type": entity,
                "commission_type": commission_type,
                "amount": amount,
                "status": "paid" if zero or entity == ENTITY_AGENCY else "ready_for_payment",
                "transaction_id": self.next_tx("TX"),
                "deleted_at": None,
                "updated_at": datetime.now(),
                "policy_uuid": record.get("forecast_policy_uuid"),
                "statement_uuid": record["uuid"],
            }
        )
        key = f"{record['agency_id']}:{line['user_uuid']}"
        if key not in self.ledgers:
            self.ledgers[key] = {
                "uuid": new_id(),
                "agency_id": record["agency_id"],
                "entity_uuid": line["user_uuid"],
                "entity_type": entity,
                "opening_balance": Decimal("0.00"),
                "deleted_at": None,
                "existing": False,
            }
        if zero or not ledger:
            return
        self.ledger_details.append(
            {
                "uuid": new_id(),
                "agency_id": record["agency_id"],
                "account_ledger_uuid": self.ledgers[key]["uuid"],
                "policy_uuid": record.get("forecast_policy_uuid"),
                "amount": abs(amount),
                "transaction_date": self.today,
                "transaction_type": commission_type,
                "credit": commission_type.upper() not in DEBIT_TYPES,
                "transaction_entity_uuid": detail_uuid,
                "description": f"Earnings for policy {record.get('policy_number')}",
                "deleted_at": None,
                "entity_uuid": line["user_uuid"],
            }
        )

    def load_reference(self):
        agency_id = self.statement["agency_id"]
        policy_ids = sorted({row["forecast_policy_uuid"] for row in self.records if row.get("forecast_policy_uuid")})
        entity_ids = sorted(
            {row["entity_uuid"] for row in self.details} | {row["entity_uuid"] for row in self.ledgers.values()}
        )
        if policy_ids:
            self.policies = self.select(
                """
                SELECT uuid::text AS uuid, agency_id, number, effective_date,
                       carrier_id::text AS carrier_id, product_id::text AS product_id
                FROM policies
                WHERE agency_id = %s AND uuid::text = ANY(%s)
                """,
                (agency_id, policy_ids),
            )
        else:
            self.policies = []
        known = {row["uuid"] for row in self.policies}
        self.policies.extend(row for row in self.local_policies if row["uuid"] not in known)
        self.contacts.extend(self.local_contacts)
        carrier_ids = [row["carrier_id"] for row in self.policies if row.get("carrier_id")]
        product_ids = [row["product_id"] for row in self.policies if row.get("product_id")]
        if carrier_ids:
            self.carriers = self.select(
                "SELECT uuid::text AS uuid, agency_id, name FROM carriers WHERE agency_id = %s AND uuid::text = ANY(%s)",
                (agency_id, carrier_ids),
            )
        if product_ids:
            self.products = self.select(
                "SELECT uuid::text AS uuid, agency_id, name FROM products WHERE agency_id = %s AND uuid::text = ANY(%s)",
                (agency_id, product_ids),
            )
        holder_ids = [row for row in policy_ids if row not in {item["uuid"] for item in self.local_policies}]
        if holder_ids:
            self.contacts.extend(
                self.select(
                    """
                    SELECT agency_id, relationship_id::text AS relationship_id, first_name, last_name
                    FROM contacts
                    WHERE agency_id = %s AND relationship_id::text = ANY(%s)
                    """,
                    (agency_id, holder_ids),
                )
            )
        if entity_ids:
            self.users = self.select(
                "SELECT uuid::text AS uuid, tenant_id, code FROM users WHERE tenant_id = %s AND uuid::text = ANY(%s)",
                (agency_id, entity_ids),
            )
            self.contacts.extend(
                self.select(
                    """
                    SELECT agency_id, relationship_id::text AS relationship_id, first_name, last_name
                    FROM contacts
                    WHERE agency_id = %s AND relationship_id::text = ANY(%s)
                    """,
                    (agency_id, entity_ids),
                )
            )
            existing = self.select(
                """
                SELECT uuid::text AS uuid, agency_id, entity_uuid::text AS entity_uuid, entity_type,
                       opening_balance, deleted_at
                FROM account_ledgers
                WHERE agency_id = %s AND entity_uuid::text = ANY(%s) AND deleted_at IS NULL
                """,
                (agency_id, entity_ids),
            )
            for row in existing:
                key = f"{row['agency_id']}:{row['entity_uuid']}"
                current = self.ledgers.get(key)
                if current and not current["existing"]:
                    for detail in self.ledger_details:
                        if detail["account_ledger_uuid"] == current["uuid"]:
                            detail["account_ledger_uuid"] = row["uuid"]
                self.ledgers[key] = {
                    **row,
                    "existing": True,
                    "opening_balance": money(row.get("opening_balance")) or Decimal("0"),
                }
            ledger_ids = [row["uuid"] for row in self.ledgers.values() if row.get("existing")]
            if ledger_ids:
                history = self.select(
                    """
                    SELECT uuid::text AS uuid, agency_id, account_ledger_uuid::text AS account_ledger_uuid,
                           policy_uuid::text AS policy_uuid, amount, transaction_date, transaction_type, credit,
                           transaction_entity_uuid::text AS transaction_entity_uuid, description, deleted_at, id
                    FROM account_ledger_details
                    WHERE agency_id = %s AND account_ledger_uuid::text = ANY(%s) AND deleted_at IS NULL
                    """,
                    (agency_id, ledger_ids),
                )
                self.ledger_details = history + self.ledger_details
        self.agencies = self.select("SELECT id, uuid::text AS uuid, name FROM agencies WHERE id = %s", (agency_id,))
        existing_rows = self.select(
            """
            SELECT comm.uuid::text AS commission_uuid, cd.uuid::text AS detail_uuid,
                   cd.entity_uuid::text AS entity_uuid, cd.entity_role, cd.commission_type AS detail_type,
                   cd.amount AS detail_amount, cd.status AS detail_status, cd.deleted_at, cd.transaction_id,
                   csr.uuid::text AS statement_uuid, csr.policy_uuid::text AS policy_uuid,
                   comm.amount AS commission_amount
            FROM commission_statement_records csr
            JOIN policy_commission_events pce
              ON pce.agency_id = csr.agency_id AND pce.statement_uuid = csr.uuid
            JOIN commissions comm
              ON comm.agency_id = pce.agency_id AND comm.policy_commission_event_uuid = pce.uuid
            LEFT JOIN commission_details cd
              ON cd.agency_id = comm.agency_id AND cd.commission_uuid = comm.uuid
            WHERE csr.file_uuid = %s::uuid AND csr.agency_id = %s
            """,
            (self.statement_uuid, agency_id),
        )
        forecasted = {row["statement_uuid"] for row in self.events}
        for row in existing_rows:
            if row["statement_uuid"] not in forecasted and row.get("detail_uuid") and row.get("deleted_at") is None:
                self.existing_details.append(row)
        # Skipped statement rows have no new details, so their payee contacts were not loaded above.
        known_contacts = {row["relationship_id"] for row in self.contacts}
        missing_entities = sorted(
            {
                row["entity_uuid"]
                for row in self.existing_details
                if row.get("entity_uuid") and row["entity_uuid"] not in known_contacts
            }
        )
        if missing_entities:
            self.contacts.extend(
                self.select(
                    """
                    SELECT agency_id, relationship_id::text AS relationship_id, first_name, last_name
                    FROM contacts
                    WHERE agency_id = %s AND relationship_id::text = ANY(%s)
                    """,
                    (agency_id, missing_entities),
                )
            )
            known_users = {row["uuid"] for row in self.users}
            missing_users = [item for item in missing_entities if item not in known_users]
            if missing_users:
                self.users.extend(
                    self.select(
                        "SELECT uuid::text AS uuid, tenant_id, code FROM users WHERE tenant_id = %s AND uuid::text = ANY(%s)",
                        (agency_id, missing_users),
                    )
                )


def write_exceptions(path, forecast):
    lines = ["# Reconciliation exceptions", ""]
    if not forecast.exceptions:
        lines.append("No statement records were left in exception for this forecast.")
    for item in forecast.exceptions:
        lines.extend(
            [
                f"## {item['policy_number'] or 'Unknown policy'} ({item['record_uuid']})",
                "",
                f"- Row order: {item['row_order']}",
                f"- Error code: {item['code']}",
                f"- Message: {item['message']}",
                f"- Match remarks: {item['match_remarks']}",
                "",
            ]
        )
    path.write_text("\n".join(lines), encoding="utf-8")


def _md_cell(value):
    if value is None:
        return ""
    return str(value).replace("|", "\\|").replace("\n", " ")


def _exception_reason(message):
    text = message or ""
    marker = "(reason="
    start = text.find(marker)
    if start < 0:
        return ""
    start += len(marker)
    end = text.find(")", start)
    if end < 0:
        return ""
    return text[start:end]


def _forecast_status_label(status):
    if status == STATUS_MATCHED:
        return "matched"
    if status == STATUS_FAILED:
        return "exception"
    if status == STATUS_EXCLUDED:
        return "excluded"
    return str(status)


def _record_detail(record):
    error = record.get("forecast_error") or {}
    if isinstance(error, dict) and error:
        return f"{error.get('code') or ''}: {error.get('message') or ''}".strip(": ")
    if record.get("forecast_policy_uuid") and not record.get("policy_exists"):
        return f"local_policy={record.get('forecast_policy_uuid')}"
    if record.get("matched_commission_uuid"):
        return f"commission={record.get('matched_commission_uuid')}"
    return ""


def _excel_cell(value):
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (int, float)):
        return value
    return str(value)


def _forecast_workbook_sheets(statement_uuid, forecast):
    matched = sum(1 for row in forecast.records if row.get("forecast_status") == STATUS_MATCHED)
    failed = sum(1 for row in forecast.records if row.get("forecast_status") == STATUS_FAILED)
    excluded = sum(1 for row in forecast.records if row.get("forecast_status") == STATUS_EXCLUDED)
    file_status = forecast.statement.get("forecast_status")
    file_label = "cleared" if file_status == FILE_STATUS_CLEARED else "waiting"
    code_counts = {}
    reason_counts = {}
    for item in forecast.exceptions:
        code = item.get("code") or "UNKNOWN"
        code_counts[code] = code_counts.get(code, 0) + 1
        reason = _exception_reason(item.get("message"))
        if reason:
            reason_counts[reason] = reason_counts.get(reason, 0) + 1
    summary = [
        ["Item", "Value"],
        ["Statement UUID", statement_uuid],
        ["Statement name", forecast.statement.get("name") or ""],
        ["Agency id", forecast.statement.get("agency_id")],
        ["File status", f"{file_status} ({file_label})"],
        ["Records", len(forecast.records)],
        ["Matched", matched],
        ["Exceptions", failed],
        ["Excluded / skipped", excluded],
        ["Local policies created", len(forecast.local_policies)],
        ["Commission events", len(forecast.commissions)],
        ["Commission details", len(forecast.details)],
    ]
    exception_codes = [["Exception code", "Count"]]
    if code_counts:
        for code, count in sorted(code_counts.items(), key=lambda item: (-item[1], item[0])):
            exception_codes.append([code, count])
    else:
        exception_codes.append(["No exceptions", ""])
    create_policy = [["Reason", "Count"]]
    if reason_counts:
        for reason, count in sorted(reason_counts.items(), key=lambda item: (-item[1], item[0])):
            create_policy.append([reason, count])
    else:
        create_policy.append(["No create-policy reason tags in exceptions", ""])
    exception_rows = [["Row", "Policy number", "Exception code", "Reason", "Message"]]
    if forecast.exceptions:
        for item in sorted(forecast.exceptions, key=lambda row: (row.get("row_order") is None, row.get("row_order") or 0)):
            exception_rows.append(
                [
                    item.get("row_order"),
                    item.get("policy_number"),
                    item.get("code"),
                    _exception_reason(item.get("message")),
                    item.get("message"),
                ]
            )
    else:
        exception_rows.append(["", "", "", "", "No exceptions"])
    local_policies = [["Policy UUID", "Product", "Effective date"]]
    if forecast.local_policies:
        for row in forecast.local_policies:
            local_policies.append([row.get("uuid"), row.get("product_id"), row.get("effective_date")])
    else:
        local_policies.append(["", "", "None"])
    row_results = [
        ["Row", "Policy number", "Gross commission", "Premium", "Chargeback", "Status", "Detail"]
    ]
    for record in sorted(forecast.records, key=lambda row: (row.get("row_order") is None, row.get("row_order") or 0)):
        row_results.append(
            [
                record.get("row_order"),
                record.get("policy_number"),
                record.get("gross"),
                record.get("premium_amount"),
                record.get("chargeback_amount"),
                _forecast_status_label(record.get("forecast_status")),
                _record_detail(record),
            ]
        )
    return [
        ("Summary", summary),
        ("Exception codes", exception_codes),
        ("Create-policy failures", create_policy),
        ("Exception rows", exception_rows),
        ("New policies", local_policies),
        ("Row results", row_results),
    ]


def write_forecast_xlsx(path, statement_uuid, forecast):
    header_font = Font(name="Calibri", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")
    alt_fill = PatternFill("solid", fgColor="D6EAF8")
    thin = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    workbook = Workbook()
    sheets = _forecast_workbook_sheets(statement_uuid, forecast)
    for index, (title, rows) in enumerate(sheets):
        sheet = workbook.active if index == 0 else workbook.create_sheet(title)
        if index == 0:
            sheet.title = title
        for row_index, row in enumerate(rows, start=1):
            for col_index, value in enumerate(row, start=1):
                cell = sheet.cell(row=row_index, column=col_index, value=_excel_cell(value))
                cell.alignment = Alignment(vertical="center", wrap_text=col_index == len(row))
                cell.border = thin
                if row_index == 1:
                    cell.font = header_font
                    cell.fill = header_fill
                elif row_index % 2 == 0:
                    cell.fill = alt_fill
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for col_index in range(1, len(rows[0]) + 1):
            width = 12
            for row in rows:
                if col_index - 1 < len(row):
                    width = max(width, min(len(str(row[col_index - 1] or "")), 60) + 2)
            sheet.column_dimensions[get_column_letter(col_index)].width = width
        sheet.row_dimensions[1].height = 18
    workbook.save(path)


def write_forecast_md(path, statement_uuid, forecast):
    matched = sum(1 for row in forecast.records if row.get("forecast_status") == STATUS_MATCHED)
    failed = sum(1 for row in forecast.records if row.get("forecast_status") == STATUS_FAILED)
    excluded = sum(1 for row in forecast.records if row.get("forecast_status") == STATUS_EXCLUDED)
    local_created = len(forecast.local_policies)
    file_status = forecast.statement.get("forecast_status")
    file_label = "cleared" if file_status == FILE_STATUS_CLEARED else "waiting"

    code_counts = {}
    reason_counts = {}
    for item in forecast.exceptions:
        code = item.get("code") or "UNKNOWN"
        code_counts[code] = code_counts.get(code, 0) + 1
        reason = _exception_reason(item.get("message"))
        if reason:
            reason_counts[reason] = reason_counts.get(reason, 0) + 1

    lines = [
        f"# Forecast log — {statement_uuid}",
        "",
        f"- Statement name: {_md_cell(forecast.statement.get('name'))}",
        f"- Agency id: {forecast.statement.get('agency_id')}",
        f"- File status: {file_status} ({file_label})",
        f"- Records: {len(forecast.records)}",
        f"- Matched: {matched}",
        f"- Exceptions: {failed}",
        f"- Excluded / skipped: {excluded}",
        f"- Local policies created: {local_created}",
        f"- Commission events: {len(forecast.commissions)}",
        f"- Commission details: {len(forecast.details)}",
        "",
        "## Exception codes",
        "",
    ]
    if code_counts:
        lines.extend(["| Code | Count |", "| --- | ---: |"])
        for code, count in sorted(code_counts.items(), key=lambda item: (-item[1], item[0])):
            lines.append(f"| {_md_cell(code)} | {count} |")
    else:
        lines.append("No exceptions.")
    lines.append("")

    lines.extend(["## Create-policy failure reasons", ""])
    if reason_counts:
        lines.extend(["| Reason | Count |", "| --- | ---: |"])
        for reason, count in sorted(reason_counts.items(), key=lambda item: (-item[1], item[0])):
            lines.append(f"| `{_md_cell(reason)}` | {count} |")
    else:
        lines.append("No create-policy reason tags in exceptions.")
    lines.append("")

    lines.extend(
        [
            "## Exception rows",
            "",
            "| Row | Policy | Code | Reason | Message |",
            "| ---: | --- | --- | --- | --- |",
        ]
    )
    if forecast.exceptions:
        for item in sorted(forecast.exceptions, key=lambda row: (row.get("row_order") is None, row.get("row_order") or 0)):
            lines.append(
                "| {row} | {policy} | {code} | {reason} | {message} |".format(
                    row=_md_cell(item.get("row_order")),
                    policy=_md_cell(item.get("policy_number")),
                    code=_md_cell(item.get("code")),
                    reason=_md_cell(_exception_reason(item.get("message"))),
                    message=_md_cell(item.get("message")),
                )
            )
    else:
        lines.append("| | | | | No exceptions |")
    lines.append("")

    lines.extend(
        [
            "## Local policies created",
            "",
            "| Policy UUID | Product | Effective |",
            "| --- | --- | --- |",
        ]
    )
    if forecast.local_policies:
        for row in forecast.local_policies:
            lines.append(
                f"| {_md_cell(row.get('uuid'))} | {_md_cell(row.get('product_id'))} | {_md_cell(row.get('effective_date'))} |"
            )
    else:
        lines.append("| | | None |")
    lines.append("")

    lines.extend(
        [
            "## Row results",
            "",
            "| Row | Policy | Gross | Premium | Chargeback | Status | Detail |",
            "| ---: | --- | ---: | ---: | ---: | --- | --- |",
        ]
    )
    for record in sorted(forecast.records, key=lambda row: (row.get("row_order") is None, row.get("row_order") or 0)):
        status = record.get("forecast_status")
        if status == STATUS_MATCHED:
            status_label = "matched"
        elif status == STATUS_FAILED:
            status_label = "exception"
        elif status == STATUS_EXCLUDED:
            status_label = "excluded"
        else:
            status_label = str(status)
        error = record.get("forecast_error") or {}
        detail = ""
        if isinstance(error, dict) and error:
            detail = f"{error.get('code') or ''}: {error.get('message') or ''}".strip(": ")
        elif record.get("forecast_policy_uuid") and not record.get("policy_exists"):
            detail = f"local_policy={record.get('forecast_policy_uuid')}"
        elif record.get("matched_commission_uuid"):
            detail = f"commission={record.get('matched_commission_uuid')}"
        lines.append(
            "| {row} | {policy} | {gross} | {premium} | {chargeback} | {status} | {detail} |".format(
                row=_md_cell(record.get("row_order")),
                policy=_md_cell(record.get("policy_number")),
                gross=_md_cell(record.get("gross")),
                premium=_md_cell(record.get("premium_amount")),
                chargeback=_md_cell(record.get("chargeback_amount")),
                status=_md_cell(status_label),
                detail=_md_cell(detail),
            )
        )
    lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def write_forecast_sql(path, forecast):
    sections = ["-- Lookup rows required by the report joins", ""]
    for row in forecast.agencies:
        sections.append(insert_sql("agencies", row))
    for row in forecast.users:
        sections.append(insert_sql("users", {"uuid": row["uuid"], "tenant_id": row["tenant_id"], "code": row["code"]}))
    for row in forecast.contacts:
        sections.append(
            insert_sql(
                "contacts",
                {
                    "agency_id": row["agency_id"],
                    "relationship_id": row["relationship_id"],
                    "first_name": row.get("first_name"),
                    "last_name": row.get("last_name"),
                },
            )
        )
    for row in forecast.carriers:
        sections.append(insert_sql("carriers", row))
    for row in forecast.products:
        sections.append(insert_sql("products", row))
    for row in forecast.policies:
        sections.append(insert_sql("policies", row))
    for row in forecast.local_commissions:
        sections.append(
            insert_sql(
                "policy_commission",
                {
                    "uuid": row["uuid"],
                    "agency_id": forecast.statement["agency_id"],
                    "policy_uuid": row["policy_uuid"],
                    "fee_amount": row.get("fee_amount"),
                    "payout_method": row.get("payout_method"),
                    "month_from": row.get("month_from"),
                    "month_to": row.get("month_to"),
                    "commission_type_text": row.get("commission_type_text"),
                    "commission_split": row.get("commission_split"),
                },
            )
        )
    sections.extend(["", "-- Forecast result", ""])
    sections.append(
        insert_sql(
            "commission_statements",
            {
                "uuid": forecast.statement_uuid,
                "agency_id": forecast.statement["agency_id"],
                "name": forecast.statement.get("name"),
                "batch_id": forecast.statement.get("batch_id"),
                "status": forecast.statement["forecast_status"],
                "stage": forecast.statement["forecast_stage"],
                "updated_at": forecast.statement.get("updated_at"),
                "carrier_uuid": forecast.statement.get("carrier_uuid"),
            },
        )
    )
    for record in forecast.records:
        sections.append(
            insert_sql(
                "commission_statement_records",
                {
                    "uuid": record["uuid"],
                    "agency_id": record["agency_id"],
                    "file_uuid": record["file_uuid"],
                    "row_order": record["row_order"],
                    "policy_number": record.get("policy_number"),
                    "policy_uuid": record.get("forecast_policy_uuid"),
                    "policy_commission_uuid": record.get("forecast_policy_commission_uuid"),
                    "status": record.get("forecast_status"),
                    "commission_type": record.get("forecast_commission_type"),
                    "error": record.get("forecast_error"),
                    "data": json_obj(record.get("data")) if record.get("data") is not None else None,
                },
            )
        )
    for event in forecast.events:
        sections.append(insert_sql("policy_commission_events", event))
    for row in forecast.commissions:
        sections.append(insert_sql("commissions", row))
    for row in forecast.details:
        sections.append(insert_sql("commission_details", row))
    for row in forecast.ledgers.values():
        sections.append(insert_sql("account_ledgers", {key: value for key, value in row.items() if key != "existing"}))
    for row in forecast.ledger_details:
        if not row.get("id"):
            sections.append(insert_sql("account_ledger_details", row))
    for row in forecast.advance_ledgers:
        sections.append(insert_sql("agent_advance_ledgers", row))
    path.write_text("\n".join(sections) + "\n", encoding="utf-8")


def index_by(rows, key):
    return {row[key]: row for row in rows if row.get(key)}


def report_rows(forecast):
    policies = index_by(forecast.policies, "uuid")
    carriers = index_by(forecast.carriers, "uuid")
    products = index_by(forecast.products, "uuid")
    users = index_by(forecast.users, "uuid")
    records = index_by(forecast.records, "uuid")
    agency = forecast.agencies[0] if forecast.agencies else {}
    contacts = defaultdict(list)
    for contact in forecast.contacts:
        contacts[contact["relationship_id"]].append(contact)

    def first_contact(rel):
        found = contacts.get(rel) or []
        return found[0] if found else {}

    detail_report = []
    combined_details = list(forecast.details)
    for row in forecast.existing_details:
        combined_details.append(
            {
                "commission_uuid": row["commission_uuid"],
                "entity_uuid": row.get("entity_uuid"),
                "entity_role": row.get("entity_role"),
                "entity_type": "",
                "commission_type": row.get("detail_type"),
                "amount": money(row.get("detail_amount")),
                "status": row.get("detail_status"),
                "transaction_id": row.get("transaction_id"),
                "policy_uuid": row.get("policy_uuid"),
                "statement_uuid": row["statement_uuid"],
                "deleted_at": None,
                "total_commission": money(row.get("commission_amount")),
            }
        )
    commissions = index_by(forecast.commissions, "uuid")
    for detail in combined_details:
        if detail.get("deleted_at") is not None:
            continue
        commission = commissions.get(detail["commission_uuid"], {})
        record = records.get(detail.get("statement_uuid") or commission.get("statement_uuid"))
        if record is None:
            continue
        policy = policies.get(detail.get("policy_uuid") or record.get("forecast_policy_uuid") or "")
        if policy and not first_contact(policy.get("uuid")):
            continue
        holder = first_contact(policy.get("uuid")) if policy else {}
        agent_contact = first_contact(detail["entity_uuid"])
        agent_name = " ".join(
            part for part in [agent_contact.get("first_name"), agent_contact.get("last_name")] if part
        ).strip()
        role_name = str(detail.get("entity_role") or detail.get("entity_type") or "").strip().lower()
        if not agent_name and role_name in {ENTITY_AGENCY, "agency"}:
            agent_name = agency.get("name") or ""
        detail_report.append(
            {
                "uploaded_at": forecast.statement.get("updated_at"),
                "transaction_id": detail.get("transaction_id"),
                "source_batch_id": forecast.statement.get("batch_id"),
                "source_file_name": forecast.statement.get("name"),
                "policy_number": record.get("policy_number"),
                "policy_effective_date": policy.get("effective_date") if policy else None,
                "policy Holder Name": " ".join(
                    part for part in [holder.get("first_name"), holder.get("last_name")] if part
                ).strip(),
                "carrier_name": carriers.get((policy or {}).get("carrier_id"), {}).get("name"),
                "product_name": products.get((policy or {}).get("product_id"), {}).get("name"),
                "Agent Name": agent_name,
                "role_name": detail.get("entity_role"),
                "paid_to_date": record.get("paid_to_date"),
                "commission_type": detail.get("commission_type"),
                "commission_amount": detail.get("amount"),
                "total_commission": commission.get("amount", detail.get("total_commission")),
                "status": detail.get("status"),
            }
        )

    dates = [forecast.today]
    for row in forecast.ledger_details:
        tx_date = row.get("transaction_date")
        dates.append(tx_date.date() if isinstance(tx_date, datetime) else tx_date or forecast.today)
    start, end = min(dates), max(dates)
    buckets = defaultdict(lambda: {name: Decimal("0") for name in ("earning", "deduction", "payment", "adjustment")})
    for row in forecast.ledger_details:
        if row.get("deleted_at") is not None:
            continue
        tx_date = row.get("transaction_date")
        tx_date = tx_date.date() if isinstance(tx_date, datetime) else tx_date
        if tx_date is None or tx_date < start or tx_date > end:
            continue
        ledger = next((item for item in forecast.ledgers.values() if item["uuid"] == row["account_ledger_uuid"]), None)
        if ledger is None:
            continue
        amount = money(row.get("amount")) or Decimal("0")
        tx_type = row.get("transaction_type")
        bucket = buckets[ledger["entity_uuid"]]
        if row.get("credit") and tx_type not in ("ADVANCE_EARNED", "ADVANCE_CONTRA"):
            bucket["earning"] += amount
        elif not row.get("credit") and tx_type != "PAYMENT":
            bucket["deduction"] += amount
        elif not row.get("credit") and tx_type == "PAYMENT":
            bucket["payment"] += amount
        elif row.get("credit") and tx_type in ("ADVANCE_EARNED", "ADVANCE_CONTRA"):
            bucket["adjustment"] += amount
    ledger_summary = []
    for entity_uuid, bucket in buckets.items():
        contact = first_contact(entity_uuid)
        ledger_summary.append(
            {
                "user_uuid": entity_uuid,
                "agent_id": users.get(entity_uuid, {}).get("code"),
                "agent_first_name": contact.get("first_name"),
                "agent_last_name": contact.get("last_name"),
                **bucket,
            }
        )
    ledger_summary.sort(key=lambda row: ((row.get("agent_last_name") or ""), (row.get("agent_first_name") or "")))

    balances = []
    for ledger in forecast.ledgers.values():
        if ledger.get("deleted_at") is not None or str(ledger.get("entity_type") or "").lower() != ENTITY_AGENT:
            continue
        pre = period = Decimal("0")
        for row in forecast.ledger_details:
            if row.get("account_ledger_uuid") != ledger["uuid"] or row.get("deleted_at") is not None:
                continue
            tx_date = row.get("transaction_date")
            tx_date = tx_date.date() if isinstance(tx_date, datetime) else tx_date
            signed = money(row.get("amount")) or Decimal("0")
            if not row.get("credit"):
                signed = -signed
            if tx_date and tx_date < start:
                pre += signed
            elif tx_date and start <= tx_date <= end:
                period += signed
        opening = (money(ledger.get("opening_balance")) or Decimal("0")) + pre
        contact = first_contact(ledger["entity_uuid"])
        balances.append(
            {
                "Agent Name": " ".join(
                    part for part in [contact.get("first_name"), contact.get("last_name")] if part
                ).strip(),
                "Agent ID": users.get(ledger["entity_uuid"], {}).get("code"),
                "opening_balance": opening,
                "closing_balance": opening + period,
            }
        )
    balances.sort(key=lambda row: row.get("Agent ID") or "")

    groups = defaultdict(lambda: defaultdict(lambda: Decimal("0")))
    for detail in combined_details:
        if detail.get("deleted_at") is not None:
            continue
        commission = commissions.get(detail["commission_uuid"], {})
        record = records.get(detail.get("statement_uuid") or commission.get("statement_uuid"))
        if record is None:
            continue
        group = groups[(forecast.statement.get("name"), record.get("policy_number"))]
        amount = money(detail.get("amount")) or Decimal("0")
        role = str(detail.get("entity_role") or "").strip().lower()
        commission_type = detail.get("commission_type")
        if commission_type == "MMP":
            group["mmp_amount"] += amount
        elif commission_type == "ARF":
            group["arf_amount"] += amount
        else:
            group["commission_split"] += amount
            if role == "agency":
                group["agency_split"] += amount
            elif role == "agent":
                group["agent_split"] += amount
            elif role == "sales leader":
                group["sales_leader_split"] += amount
            elif role in ("subagent", "sub agent", "sub-agent"):
                group["subagent_split"] += amount
        data = json_obj(record.get("data"))
        if isinstance(data, dict):
            earned = money(data.get("earned_commission"))
            gross = money(data.get("gross_commission"))
            if earned is not None:
                group["earned_commission"] = earned
            if gross is not None:
                group["gross_commission"] = gross
    file_status = forecast.statement["forecast_status"]
    split_report = [
        {
            "source_file_name": name,
            "policy_number": policy,
            "file_status": file_status,
            **{column: values[column] for column in SPLIT_COLUMNS},
        }
        for (name, policy), values in sorted(groups.items(), key=lambda item: (item[0][0] or "", item[0][1] or ""))
    ]
    return detail_report, ledger_summary, balances, split_report, start, end


def write_csv(path, rows):
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    fieldnames = []
    seen = set()
    for row in rows:
        for key in row:
            if key not in seen:
                fieldnames.append(key)
                seen.add(key)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, restval="")
        writer.writeheader()
        writer.writerows(rows)


def sanitize_folder_component(value):
    text = str(value or "").strip()
    if not text:
        return ""
    for char in ('/', '\\', '\0', ':', '*', '?', '"', '<', '>', '|'):
        text = text.replace(char, "_")
    return text.strip(" .")


def forecast_output_dir(statement_uuid, statement_name=None):
    safe_name = sanitize_folder_component(statement_name)
    folder = f"{statement_uuid} - {safe_name}" if safe_name else str(statement_uuid)
    return Path("forecast-out") / folder


def write_forecast_output(statement_uuid, forecast, out=None):
    if out is None:
        out = forecast_output_dir(statement_uuid, forecast.statement.get("name"))
    out.mkdir(parents=True, exist_ok=True)
    write_exceptions(out / "reconciliation_exceptions.md", forecast)
    write_forecast_md(out / "forecast.md", statement_uuid, forecast)
    write_forecast_xlsx(out / "forecast.xlsx", statement_uuid, forecast)
    write_forecast_sql(out / "forecast.sql", forecast)
    details, ledger, balances, splits, start, end = report_rows(forecast)
    if forecast.statement["forecast_status"] != FILE_STATUS_CLEARED:
        logger.info(
            "Split report includes processed lines. The published query returns no rows while statement status is %s because %s record(s) are still in exception.",
            forecast.statement["forecast_status"],
            len(forecast.exceptions),
        )
    write_csv(out / "commission_detail_report.csv", details)
    write_csv(out / "ledger_summary.csv", ledger)
    write_csv(out / "agent_balances.csv", balances)
    write_csv(out / "commission_split_report.csv", splits)
    (out / "summary.json").write_text(
        json.dumps(
            {
                "statement_uuid": statement_uuid,
                "agency_id": forecast.statement["agency_id"],
                "forecast_file_status": forecast.statement["forecast_status"],
                "exceptions": len(forecast.exceptions),
                "commissions": len(forecast.commissions),
                "commission_details": len(forecast.details),
                "report_date_from": start.isoformat(),
                "report_date_to": end.isoformat(),
                "api_base_url_mutating_calls": 0,
                "environment_database_writes": 0,
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    logger.info("Wrote forecast files to %s", out.resolve())
    run_apl_cross_validation(out, forecast.statement.get("name"))


def resolve_apl_statement(statement_name):
    override = os.environ.get("APL_STATEMENT_PATH", "").strip()
    if override:
        path = Path(override)
        return path if path.is_file() else None
    directory = os.environ.get("APL_STATEMENT_DIR", "").strip()
    if not directory or not statement_name:
        return None
    root = Path(directory)
    direct = root / statement_name
    if direct.is_file():
        return direct
    target = statement_name.lower()
    if not root.is_dir():
        return None
    for candidate in root.rglob("*"):
        if candidate.is_file() and candidate.name.lower() == target:
            return candidate
    return None


def load_apl_validator():
    script = Path(__file__).resolve().parents[1] / "scripts" / "apl_icm_cross_validation" / "validate_apl_icm.py"
    spec = importlib.util.spec_from_file_location("validate_apl_icm", script)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def resolve_apl_validation_mode():
    raw = os.environ.get("APL_VALIDATION_MODE", "apl").strip() or "apl"
    aliases = {
        "1": "apl",
        "apl": "apl",
        "apl_to_icm": "apl",
        "2": "icm",
        "icm": "icm",
        "icm_to_apl": "icm",
        "parallel": "icm",
    }
    mode = aliases.get(raw.lower(), raw.lower())
    if mode not in ("apl", "icm"):
        logger.warning("Unknown APL_VALIDATION_MODE=%s; using apl", raw)
        return "apl"
    return mode


def run_apl_cross_validation(out, statement_name):
    detail_csv = out / "commission_detail_report.csv"
    if not detail_csv.is_file() or detail_csv.stat().st_size == 0:
        logger.info("APL cross-validation skipped: commission_detail_report.csv has no rows")
        return
    apl_file = resolve_apl_statement(statement_name)
    if apl_file is None:
        logger.info(
            "APL cross-validation skipped for %s. Set APL_STATEMENT_PATH or APL_STATEMENT_DIR to the carrier statement.",
            statement_name,
        )
        return
    mode = resolve_apl_validation_mode()
    validation_dir = out / "apl_icm_validation"
    validation_dir.mkdir(parents=True, exist_ok=True)
    logger.info("Running APL cross-validation: mode=%s apl=%s parallel=%s", mode, apl_file, detail_csv)
    try:
        validator = load_apl_validator()
        validator.validate_data(
            str(apl_file),
            str(detail_csv),
            str(validation_dir / "Validation_Output.csv"),
            str(validation_dir / "Validation_Output.xlsx"),
            mode=mode,
        )
    except Exception as exc:
        logger.exception("APL cross-validation failed: %s", exc)
        return
    logger.info("APL cross-validation written to %s", validation_dir)


STATEMENT_NAME_SQL = """
SELECT name
FROM commission_statements
WHERE uuid = %s::uuid
"""


def resolve_statement_name(conn, statement_uuid, statement_name=None):
    if statement_name:
        return statement_name
    rows = fetch_all(conn, STATEMENT_NAME_SQL, (statement_uuid,))
    return rows[0]["name"] if rows else None


def run_statement(conn, statement_uuid, auth, api_base_url, statement_name=None):
    name = resolve_statement_name(conn, statement_uuid, statement_name)
    out = forecast_output_dir(statement_uuid, name)
    log_path = out / "forecast.log"
    handler = attach_forecast_log(log_path)
    try:
        forecast = Forecast(conn, statement_uuid, auth, api_base_url)
        forecast.reconcile()
        forecast.process_commissions()
        forecast.load_reference()
        write_forecast_output(statement_uuid, forecast, out=out)
    finally:
        detach_forecast_log(handler)


def parse_agency_id(argv):
    agency_id = None
    remaining = []
    index = 0
    while index < len(argv):
        token = argv[index]
        if token in ("--agency_id", "--agency-id"):
            if index + 1 >= len(argv):
                raise ForecastError("--agency_id requires a number")
            agency_id = int(argv[index + 1])
            index += 2
            continue
        remaining.append(token)
        index += 1
    return remaining, agency_id


def statement_uuids_from_args(tokens):
    statement_uuids = []
    for token in tokens:
        for part in token.split(","):
            value = part.strip()
            if not value:
                continue
            try:
                statement_uuids.append(str(uuid.UUID(value)))
            except ValueError as exc:
                raise ForecastError(
                    f"Not a commission statement UUID: {value}. "
                    "Pass one or more UUIDs, or review --agency_id <id>"
                ) from exc
    if not statement_uuids:
        raise ForecastError("Pass one or more commission statement UUIDs, or review --agency_id <id>")
    return statement_uuids


def main():
    configure_logging()
    try:
        argv, agency_id = parse_agency_id(sys.argv[1:])
    except (ForecastError, ValueError) as exc:
        logger.error("%s", exc)
        return 2
    usage = (
        "Usage: review_to_reconcile_file_forecast.py <preprod|prod> "
        "<commission_statement_uuid> [<uuid> ...] | review [--agency_id <id>]"
    )
    if len(argv) < 2 or argv[0] not in ("preprod", "prod"):
        logger.error(usage)
        return 2
    environment = argv[0]
    targets = argv[1:]
    review_all = targets == ["review"]
    if "review" in targets and not review_all:
        logger.error("Pass review alone with --agency_id, or pass statement UUIDs without review")
        return 2
    if review_all and agency_id is None:
        logger.error("Review mode requires --agency_id <id>")
        return 2
    statement_uuids = []
    if not review_all:
        try:
            statement_uuids = statement_uuids_from_args(targets)
        except ForecastError as exc:
            logger.error("%s", exc)
            return 2
    logger.info(
        "Starting %s forecast for %s%s",
        environment,
        "review files" if review_all else f"{len(statement_uuids)} file(s)",
        f" agency_id={agency_id}" if agency_id is not None else "",
    )
    try:
        load_env_file(environment)
    except ForecastError as exc:
        logger.error("%s", exc)
        return 1
    try:
        auth = {"token": obtain_token()}
        db_url, db_password = resolve_database()
        conn = connect_readonly(db_url, db_password)
        api_base = require_env("PIEQ_API_BASE_URL")
        try:
            if review_all:
                files = fetch_all(conn, REVIEW_FILES_SQL, (REVIEW_STAGE, agency_id))
                logger.info(
                    "Found %s commission statement(s) in Review (stage %s) for agency_id=%s",
                    len(files),
                    REVIEW_STAGE,
                    agency_id,
                )
                if not files:
                    return 0
                failed = 0
                for index, row in enumerate(files, start=1):
                    logger.info(
                        "Forecasting review file %s/%s uuid=%s name=%s",
                        index,
                        len(files),
                        row["uuid"],
                        row.get("name"),
                    )
                    try:
                        run_statement(conn, row["uuid"], auth, api_base, statement_name=row.get("name"))
                    except Exception as exc:
                        failed += 1
                        logger.exception("Forecast failed for %s: %s", row["uuid"], exc)
                logger.info("Review forecast finished: files=%s failed=%s", len(files), failed)
                return 1 if failed else 0
            failed = 0
            for index, statement_uuid in enumerate(statement_uuids, start=1):
                logger.info(
                    "Forecasting file %s/%s uuid=%s",
                    index,
                    len(statement_uuids),
                    statement_uuid,
                )
                try:
                    run_statement(conn, statement_uuid, auth, api_base)
                except Exception as exc:
                    failed += 1
                    logger.exception("Forecast failed for %s: %s", statement_uuid, exc)
            logger.info("Forecast finished: files=%s failed=%s", len(statement_uuids), failed)
            return 1 if failed else 0
        finally:
            conn.close()
    except ForecastError as exc:
        logger.error("%s", exc)
        return 1
    except Exception as exc:
        logger.exception("Forecast failed: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
