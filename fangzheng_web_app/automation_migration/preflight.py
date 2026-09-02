from __future__ import annotations

from typing import Any


BOOLEAN_REQUIREMENTS = (
    "code_review_complete", "automation_tests_passed", "full_data_verification_passed",
    "postgres_backup_restore_rehearsed", "sqlite_rollback_rehearsed", "file_downloads_verified",
    "business_acceptance_passed", "monitoring_alerts_configured",
    "switch_owner_assigned", "review_owner_assigned", "rollback_owner_assigned",
)


def evaluate_preflight(evidence: dict[str, Any]) -> dict[str, Any]:
    blockers = [name for name in BOOLEAN_REQUIREMENTS if evidence.get(name) is not True]
    if int(evidence.get("shadow_observation_days", 0)) < 7:
        blockers.append("shadow_observation_days>=7")
    if int(evidence.get("unexplained_shadow_differences", -1)) != 0:
        blockers.append("unexplained_shadow_differences=0")
    if int(evidence.get("outbox_pending", -1)) != 0:
        blockers.append("outbox_pending=0")
    if not evidence.get("performance_thresholds_approved"):
        blockers.append("performance_thresholds_approved")
    if not evidence.get("performance_test_passed"):
        blockers.append("performance_test_passed")
    return {"passed": not blockers, "blockers": blockers}
