"""Compatibility aliases for the unified confirmation policy.

New runtime code must import ``transcode_confirmation_policy`` directly.
"""

from .transcode_confirmation_policy import (
    DEFAULT_CONFIRMATION_RULE_PATH as DEFAULT_RISK_PATH,
    load_confirmation_policy_rules as load_pending_formal_risks,
    match_confirmation_policy_rules as match_pending_formal_risks,
)


def match_pending_formal_risk(customer: str, *texts: str):
    matches = match_pending_formal_risks(customer, *texts)
    return matches[0] if matches else None
