from __future__ import annotations

from flask import Flask

from fangzheng_web_app.connector_api import bp


def _app(monkeypatch) -> Flask:
    monkeypatch.setenv("CONNECTOR_API_TOKEN", "connector-test-token")
    app = Flask(__name__)
    app.testing = True
    app.register_blueprint(bp)
    return app


def _headers(employee_id: str = "23582") -> dict[str, str]:
    return {"Authorization": "Bearer connector-test-token", "X-Nanya-Employee-Id": employee_id}


def _mcp(client, method: str, params: dict | None = None):
    return client.post("/api/connector/v1/mcp", headers=_headers(), json={"jsonrpc": "2.0", "id": 1, "method": method, "params": params or {}})


def test_mcp_tools_list_is_machine_json_without_html(monkeypatch):
    client = _app(monkeypatch).test_client()
    response = _mcp(client, "tools/list")
    assert response.status_code == 200 and response.is_json
    assert {tool["name"] for tool in response.get_json()["result"]["tools"]} == {
        "marketing.mail.sync_orders", "marketing.job.get", "marketing.order.list_cases", "marketing.order.get_case",
        "marketing.material.query", "marketing.order.prepare_entry", "marketing.order.reply_draft", "marketing.order.submit_entry",
    }
    assert b"<html" not in response.data.lower()


def test_identity_and_token_are_required(monkeypatch):
    client = _app(monkeypatch).test_client()
    assert client.get("/api/connector/v1/health").status_code == 403
    assert client.get("/api/connector/v1/health", headers=_headers("99999")).status_code == 403


def test_health_redacts_mail_secrets_and_uses_23582(monkeypatch):
    from fangzheng_web_app import connector_api
    monkeypatch.setattr(connector_api.mail_store, "list_accounts", lambda **_: [{"id": 3, "enabled": 1, "imap_host": "imap.example", "auth_code_ciphertext": "do-not-leak", "smtp_auth_code_ciphertext": "do-not-leak"}])
    monkeypatch.setattr(connector_api.mail_store, "smtp_public_config", lambda _: {"configured": True, "enabled": False})
    monkeypatch.setattr(connector_api, "get_interface_config", lambda _: {"enabled": True, "mode": "mock"})
    response = _app(monkeypatch).test_client().get("/api/connector/v1/health", headers=_headers())
    assert response.status_code == 200 and response.get_json()["employee_id"] == "23582"
    assert "do-not-leak" not in response.get_data(as_text=True)


def test_mcp_sync_calls_existing_queue_with_employee_context(monkeypatch):
    from fangzheng_web_app import connector_api
    monkeypatch.setattr(connector_api, "_business_account", lambda *_: {"id": 8})
    captured = {}
    def queue(account_id, **kwargs):
        captured.update({"account_id": account_id, **kwargs})
        return {"fetch_task_id": 88, "message": "queued"}
    monkeypatch.setattr(connector_api.mail_fetch_service, "queue_latest_order_mails", queue)
    response = _mcp(_app(monkeypatch).test_client(), "tools/call", {"name": "marketing.mail.sync_orders", "arguments": {"lookback_days": 2}})
    assert response.status_code == 200
    assert captured == {"account_id": 8, "created_by": "23582", "owner_employee_id": "23582", "lookback_days": 2, "limit": None}
    assert '"id": 88' in response.get_json()["result"]["content"][0]["text"]
