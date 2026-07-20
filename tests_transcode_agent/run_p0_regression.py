from __future__ import annotations

from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parents[1]

COMMANDS = [
    [
        sys.executable,
        "-m",
        "py_compile",
        "fangzheng_web_app/transcode_engine.py",
        "fangzheng_web_app/transcode_evidence_scoring.py",
        "fangzheng_web_app/transcode_evidence_model.py",
        "fangzheng_web_app/transcode_agent_rules.py",
        "fangzheng_web_app/transcode_agent_service.py",
        "fangzheng_web_app/transcode_semantic_service.py",
        "fangzheng_web_app/transcode_semantic_rule_maintenance.py",
        "fangzheng_web_app/transcode_semantic_rule_compiler.py",
        "fangzheng_web_app/transcode_semantic_rule_finalizer.py",
        "fangzheng_web_app/transcode_semantic_rules.py",
        "model_skills/customer-special-rule-maintenance/scripts/finalize_semantic_rule_rows.py",
        "fangzheng_web_app/transcode_semantic_shadow.py",
        "fangzheng_web_app/transcode_semantic_overrides.py",
        "fangzheng_web_app/transcode_pending_risks.py",
        "tests_transcode_agent/test_confirmed_draft_rules.py",
        "tests_transcode_agent/smoke_confirmed_draft_rule_version.py",
        "tests_transcode_agent/smoke_agent_rule_application.py",
        "tests_transcode_agent/smoke_agent_batch_sample.py",
        "tests_transcode_agent/smoke_agent_active_version_runtime.py",
        "tests_transcode_agent/smoke_agent_batch_job_runtime.py",
        "tests_transcode_agent/smoke_agent_mapping_tables.py",
        "tests_transcode_agent/smoke_agent_size_mapping_application.py",
        "tests_transcode_agent/smoke_agent_thickness_mapping_application.py",
        "tests_transcode_agent/smoke_agent_manual_upload_sample.py",
        "tests_transcode_agent/smoke_agent_confirmed_feedback_fixes.py",
        "tests_transcode_agent/smoke_agent_semantic_model_foundation.py",
        "tests_transcode_agent/smoke_agent_evidence_scoring.py",
        "tests_transcode_agent/smoke_agent_evidence_model.py",
        "tests_transcode_agent/smoke_agent_semantic_rule_maintenance.py",
        "tests_transcode_agent/smoke_agent_semantic_rule_compiler.py",
        "tests_transcode_agent/smoke_agent_semantic_rule_finalizer.py",
        "tests_transcode_agent/smoke_agent_semantic_rule_version.py",
        "tests_transcode_agent/smoke_agent_semantic_shadow.py",
        "tests_transcode_agent/smoke_agent_semantic_overrides.py",
        "tests_transcode_agent/smoke_agent_pending_risks.py",
        "tests_transcode_agent/smoke_agent_bundled_defaults.py",
        "tests_transcode_agent/smoke_clean_checkout_startup.py",
        "tests_transcode_agent/smoke_agent_semantic_shadow_batch.py",
    ],
    [sys.executable, "tests_transcode_agent/smoke_clean_checkout_startup.py"],
    [sys.executable, "tests_transcode_agent/smoke_confirmed_draft_rule_version.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_rule_application.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_batch_sample.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_active_version_runtime.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_batch_job_runtime.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_mapping_tables.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_size_mapping_application.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_thickness_mapping_application.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_manual_upload_sample.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_confirmed_feedback_fixes.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_semantic_model_foundation.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_evidence_scoring.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_evidence_model.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_semantic_rule_maintenance.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_semantic_rule_compiler.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_semantic_rule_finalizer.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_semantic_rule_version.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_semantic_shadow.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_semantic_overrides.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_pending_risks.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_bundled_defaults.py"],
    [sys.executable, "tests_transcode_agent/smoke_agent_semantic_shadow_batch.py"],
]


def main() -> None:
    for command in COMMANDS:
        print(f"$ {' '.join(command)}", flush=True)
        subprocess.run(command, cwd=ROOT, check=True)
    print("P0 regression passed")


if __name__ == "__main__":
    main()
