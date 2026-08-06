from __future__ import annotations

from fangzheng_web_app import create_app, db
from fangzheng_web_app.transcode_order_semantic_model import (
    load_order_semantic_runtime,
)
from fangzheng_web_app.transcode_model_config import update_user_model_config


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
        session["employee_id"] = "model-config-tester"
    return client


def test_saved_key_is_masked_by_default_and_can_be_toggled(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)
    update_user_model_config(
        "model-config-tester",
        enabled=False,
        base_url="https://api.deepseek.com",
        api_key="saved-test-key",
        model="deepseek-v4-pro",
    )

    response = client.get("/features/transcode-agent/model-config")

    assert response.status_code == 200
    body = response.get_data(as_text=True)
    assert 'id="model-api-key"' in body
    assert 'type="password"' in body
    assert 'value="saved-test-key"' in body
    assert 'id="toggle-model-api-key"' in body
    assert "显示 API Key" in body


def test_save_redirects_to_transcode_agent(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)

    response = client.post(
        "/features/transcode-agent/model-config",
        data={
            "action": "save",
            "base_url": "https://api.deepseek.com",
            "model": "deepseek-v4-pro",
            "api_key": "saved-test-key",
        },
    )

    assert response.status_code == 302
    assert response.headers["Location"].endswith("/features/transcode-agent")


def test_save_error_stays_on_model_config_page(monkeypatch, tmp_path):
    client = _client(monkeypatch, tmp_path)

    response = client.post(
        "/features/transcode-agent/model-config",
        data={
            "action": "save",
            "base_url": "http://insecure.example.com",
            "model": "deepseek-v4-pro",
            "api_key": "saved-test-key",
        },
    )

    assert response.status_code == 302
    assert response.headers["Location"].endswith(
        "/features/transcode-agent/model-config"
    )


def test_user_enabled_config_enables_runtime(monkeypatch, tmp_path):
    _client(monkeypatch, tmp_path)
    update_user_model_config(
        "model-config-tester",
        enabled=True,
        base_url="https://api.deepseek.com",
        api_key="user-test-key",
        model="deepseek-v4-pro",
    )

    runtime = load_order_semantic_runtime("model-config-tester")

    assert runtime.mode == "active"
    assert runtime.client is not None
    assert runtime.load_error == ""

    update_user_model_config(
        "model-config-tester",
        enabled=False,
        base_url="https://api.deepseek.com",
        api_key="user-test-key",
        model="deepseek-v4-pro",
    )
    runtime = load_order_semantic_runtime("model-config-tester")
    assert runtime.mode == "off"
