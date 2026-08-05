from __future__ import annotations

from pathlib import Path

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
    assert "当前生效范围" in overview.get_data(as_text=True)
    assert "当前基础版本" in overview.get_data(as_text=True)
    assert "Agent 版本" in overview.get_data(as_text=True)
    base = client.get(
        "/admin/transcode-rule-center?section=base&business_category=胶系"
    )
    body = base.get_data(as_text=True)
    assert "当前实际生效的基础规则" in body
    assert "data-source-order=\"老表|新表|额外正式补充\"" in body
    assert ">编辑</a>" in body
    assert "新增胶系映射" in body
    assert "新增直接对应" not in body
    assert "新增条件规则" not in body
    assert "正式业务映射表" in body
    assert "transcode_rules" not in body
    assert "系统内置" not in body
    assert "<td><a href=" not in body
    assert "data-rule-list" in body
    assert "rule-list-controls.js" in body
    assert "基础规则维护方式" not in body
    assert "标准映射</a>" not in body
    assert "客户与算法规则</a>" not in body
    assert ">来源<" not in body

    mapping = client.get(
        "/admin/transcode-rule-center?section=base&business_category=基板尺寸"
    )
    mapping_body = mapping.get_data(as_text=True)
    assert mapping.status_code == 200
    assert "相同标准结果合并展示" not in mapping_body
    assert "查看原始规则明细" not in mapping_body
    assert ">来源<" not in mapping_body

    semantic = client.get("/admin/transcode-rule-center?section=semantic")
    assert semantic.status_code == 302
    assert semantic.headers["Location"].endswith("/admin/transcode-rule-center?section=customer")


def test_customer_rule_archive_uses_shared_search_and_paging(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)

    response = client.get("/admin/transcode-agent-customer-rules?rule_kind=all")

    assert response.status_code == 200
    body = response.get_data(as_text=True)
    assert "客户特殊规则维护" in body
    assert "data-rule-list" in body
    assert "rule-list-controls.js" in body
    assert "确定性特殊规则" not in body
    assert "订单备注语义规则</a>" not in body
    assert "生效规则" in body
    assert "待完善" in body
    assert "全部资料" not in body
    assert "当前查看" not in body
    assert "参考资料" not in body
    assert "历史与修改记录" in body
    assert "<th>操作</th>" in body
    assert "cr-edit-action" in body


def test_customer_rule_form_keeps_agent_override_sync_controls(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)

    response = client.get(
        "/admin/transcode-agent-customer-rules?new=1&rule_type=deterministic"
    )

    assert response.status_code == 200
    body = response.get_data(as_text=True)
    assert 'id="cr-agent-override-field"' in body
    assert 'data-business="铜箔类型+印字/非印字"' in body
    assert "refreshAgentOverride" in body


def test_transcode_agent_page_hides_legacy_rule_display(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)

    response = client.get("/features/transcode-agent")

    assert response.status_code == 200
    body = response.get_data(as_text=True)
    assert "规则配置" in body
    assert "当前规则" not in body
    assert "Agent规则说明" not in body
    assert "旧版转码" not in body


def test_confirmation_center_buttons_and_auto_basis_are_business_friendly():
    template = (
        Path(__file__).resolve().parents[1]
        / "templates"
        / "transcode_agent_confirmation.html"
    ).read_text(encoding="utf-8")

    assert "确认这行并写正式码" in template
    assert "这行先不处理，保留待确认" in template
    assert "业务确认：" in template
    assert "refreshAutoBasis" in template


def test_legacy_agent_rule_doc_redirects_to_rule_center(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)

    response = client.get("/rules-docs/transcode_agent")

    assert response.status_code == 302
    assert response.headers["Location"].endswith("/admin/transcode-rule-center")


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
