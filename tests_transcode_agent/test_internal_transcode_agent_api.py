from __future__ import annotations

from fangzheng_web_app import create_app
from fangzheng_web_app import routes


TOKEN = "test-internal-token"


def _client(monkeypatch, calculator):
    monkeypatch.setenv("TRANSCODE_AGENT_INTERNAL_TOKEN", TOKEN)
    monkeypatch.setattr(routes, "calculate_transcode_agent_quote", calculator)
    app = create_app()
    app.config.update(TESTING=True)
    return app.test_client()


def test_internal_quote_success(monkeypatch):
    captured = {}

    def fake_calculator(spec, **kwargs):
        captured.update(spec=spec, **kwargs)
        return {
            "status": "成功",
            "result": "2B00800137004900YWA1T*",
            "candidate_code": "2B00800137004900YWA1T*",
            "pending_code": "",
            "confidence": 100,
            "note": "全部关键字段达到100分",
            "error": "",
            "field_evidence": [{"field": "胶系", "score": 100}],
            "rule_version": "base-v1",
            "agent_rule_version": "agent-v1",
        }

    monkeypatch.setenv("TRANSCODE_AGENT_INTERNAL_EMPLOYEE_ID", "inventory-service")
    client = _client(monkeypatch, fake_calculator)
    response = client.post(
        "/api/internal/transcode-agent/quote",
        headers={"X-Internal-Token": TOKEN},
        json={
            "customer_code": "103901",
            "customer": "广东依顿",
            "spec": 'NY2150 0.8mm 1/1 37*49" HTE 含铜',
            "order_remark": "下汽车板",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["status"] == "成功"
    assert payload["formal_code"] == "2B00800137004900YWA1T*"
    assert payload["candidate_code"] == payload["formal_code"]
    assert payload["pending_code"] == ""
    assert payload["confidence"] == 100
    assert payload["requires_manual_completion"] is True
    assert payload["incomplete_fields"] == ["结构码"]
    assert payload["aps_query_ready"] is False
    assert captured == {
        "spec": 'NY2150 0.8mm 1/1 37*49" HTE 含铜',
        "customer": "广东依顿",
        "customer_code": "103901",
        "order_remark": "下汽车板",
        "employee_id": "inventory-service",
    }


def test_internal_quote_pending_never_returns_formal_code(monkeypatch):
    def fake_calculator(spec, **kwargs):
        return {
            "status": "待确认",
            "result": "SHOULD-NOT-BE-FORMAL",
            "candidate_code": "2B00800137004900YWA1T*",
            "pending_code": "",
            "confidence": 98,
            "note": "基板级别存在规则冲突",
            "error": "请人工确认基板级别",
            "field_evidence": [{"field": "基板级别", "score": 98}],
            "rule_version": "base-v1",
            "agent_rule_version": "agent-v1",
        }

    client = _client(monkeypatch, fake_calculator)
    response = client.post(
        "/api/internal/transcode-agent/quote",
        headers={"X-Internal-Token": TOKEN},
        json={"spec": "NY2150 0.8mm 1/1 37*49 HTE"},
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["status"] == "待确认"
    assert payload["formal_code"] == ""
    assert payload["candidate_code"] == "2B00800137004900YWA1T*"
    assert payload["pending_code"] == payload["candidate_code"]
    assert payload["reason"] == "请人工确认基板级别"
    assert payload["requires_manual_completion"] is True
    assert payload["incomplete_fields"] == ["结构码"]
    assert payload["aps_query_ready"] is False


def test_internal_quote_complete_structure_is_ready_for_aps(monkeypatch):
    def fake_calculator(spec, **kwargs):
        return {
            "status": "成功",
            "result": "2B00300HH41004900YWA1CA",
            "candidate_code": "2B00300HH41004900YWA1CA",
            "confidence": 100,
            "note": "全部关键字段达到100分",
            "error": "",
            "field_evidence": [
                {
                    "field_key": "structure",
                    "field": "结构码",
                    "code": "A",
                    "hit_type": "Agent规则覆盖",
                    "gate": False,
                }
            ],
            "rule_version": "base-v1",
            "agent_rule_version": "agent-v1",
        }

    client = _client(monkeypatch, fake_calculator)
    response = client.post(
        "/api/internal/transcode-agent/quote",
        headers={"X-Internal-Token": TOKEN},
        json={"spec": "NY2150 0.3mm H/H 41x49 HTE 芯厚"},
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["formal_code"] == "2B00300HH41004900YWA1CA"
    assert payload["requires_manual_completion"] is False
    assert payload["incomplete_fields"] == []
    assert payload["aps_query_ready"] is True


def test_internal_quote_business_failure(monkeypatch):
    def fake_calculator(spec, **kwargs):
        return {
            "status": "失败",
            "result": None,
            "candidate_code": "",
            "confidence": 0,
            "note": "无法识别胶系型号",
            "error": "无法识别胶系型号",
            "field_evidence": [],
            "rule_version": "base-v1",
            "agent_rule_version": "agent-v1",
        }

    client = _client(monkeypatch, fake_calculator)
    response = client.post(
        "/api/internal/transcode-agent/quote",
        headers={"X-Internal-Token": TOKEN},
        json={"spec": "无法识别的规格"},
    )

    assert response.status_code == 422
    payload = response.get_json()
    assert payload["status"] == "失败"
    assert payload["formal_code"] == ""
    assert payload["error"] == "无法识别胶系型号"


def test_internal_quote_rejects_when_token_is_not_configured(monkeypatch):
    # 显式空值代表部署环境禁用，不应被本地配置回填。
    monkeypatch.setenv("TRANSCODE_AGENT_INTERNAL_TOKEN", "")
    app = create_app()
    app.config.update(TESTING=True)

    response = app.test_client().post(
        "/api/internal/transcode-agent/quote",
        json={"spec": "NY2150 0.8mm 1/1 37*49 HTE"},
    )

    assert response.status_code == 503
    assert "未配置" in response.get_json()["error"]


def test_internal_quote_rejects_missing_and_wrong_tokens(monkeypatch):
    client = _client(monkeypatch, lambda spec, **kwargs: {})

    missing = client.post(
        "/api/internal/transcode-agent/quote",
        json={"spec": "NY2150 0.8mm 1/1 37*49 HTE"},
    )
    wrong = client.post(
        "/api/internal/transcode-agent/quote",
        headers={"X-Internal-Token": "wrong-token"},
        json={"spec": "NY2150 0.8mm 1/1 37*49 HTE"},
    )

    assert missing.status_code == 401
    assert wrong.status_code == 401
    assert missing.get_json()["error"] == "内部服务 Token 无效"
    assert wrong.get_json()["error"] == "内部服务 Token 无效"
