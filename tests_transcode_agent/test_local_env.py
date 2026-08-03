from __future__ import annotations

from fangzheng_web_app.local_env import load_local_env


def test_load_local_env_fills_missing_values_without_overwriting(monkeypatch, tmp_path):
    path = tmp_path / "local.env"
    path.write_text(
        "# local only\nFIRST=value-1\nSECOND='value 2'\n",
        encoding="utf-8",
    )
    monkeypatch.delenv("FIRST", raising=False)
    monkeypatch.setenv("SECOND", "explicit")

    load_local_env(path)

    assert __import__("os").environ["FIRST"] == "value-1"
    assert __import__("os").environ["SECOND"] == "explicit"
