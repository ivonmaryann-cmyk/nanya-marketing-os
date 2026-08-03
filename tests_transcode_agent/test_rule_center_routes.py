from __future__ import annotations

from fangzheng_web_app import create_app, db


def _client(monkeypatch, tmp_path):
    monkeypatch.setattr(db, "DATABASE_PATH", tmp_path / "app.db")
    monkeypatch.setattr(
        "fangzheng_web_app.transcode_rule_center.BACKUP_DIR",
        tmp_path / "backups",
    )
    app = create_app()
    app.config.update(TESTING=True)
    client = app.test_client()
    with client.session_transaction() as session:
        session["employee_id"] = "rule-center-tester"
    return client


def test_rule_center_requires_login(monkeypatch, tmp_path):
    monkeypatch.setattr(db, "DATABASE_PATH", tmp_path / "app.db")
    app = create_app()
    app.config.update(TESTING=True)

    response = app.test_client().get("/admin/transcode-rule-center")

    assert response.status_code == 302
    assert response.headers["Location"].endswith("/login")


def test_all_rule_center_sections_render(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)

    for section in ("overview", "base", "scoring", "backups", "history"):
        response = client.get(f"/admin/transcode-rule-center?section={section}")
        assert response.status_code == 200
        assert "营销转码 Agent 规则配置" in response.get_data(as_text=True)

    overview = client.get("/admin/transcode-rule-center?section=overview")
    assert "规则覆盖检查" in overview.get_data(as_text=True)
    assets = client.get(
        "/admin/transcode-rule-center?section=base&base_view=assets&asset_group=Agent胶系主表"
    )
    body = assets.get_data(as_text=True)
    assert "客户与算法规则" in body
    assert "最新版胶系主表" in body
    assert "> × </a>" not in body


def test_invalid_score_post_is_rejected_without_server_error(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)

    response = client.post(
        "/admin/transcode-rule-center?section=scoring",
        data={
            "action": "save_score",
            "semantic_supported_score": "100",
            "model_supported_score": "95",
            "ambiguous_score": "80",
            "missing_evidence_score": "60",
            "contradicted_score": "0",
        },
    )

    assert response.status_code == 200
    body = response.get_data(as_text=True)
    assert "非确定性评分必须填写0到99之间" in body


def test_corrupt_backup_restore_is_reported_without_server_error(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)
    backup_dir = tmp_path / "backups"
    corrupt = backup_dir / "transcode-rules-2026-08-02-corrupt.json"
    corrupt.write_text("not-json", encoding="utf-8")

    response = client.post(
        "/admin/transcode-rule-center?section=backups",
        data={"action": "restore_backup", "backup_name": corrupt.name},
    )

    assert response.status_code == 200
    assert "规则配置失败" in response.get_data(as_text=True)
