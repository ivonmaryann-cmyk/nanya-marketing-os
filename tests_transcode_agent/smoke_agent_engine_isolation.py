from __future__ import annotations

from fangzheng_web_app import transcode_agent_service, transcode_service


def main() -> None:
    assert transcode_service.TRANSCODE_MODULE_NAME == "fangzheng_web_app.transcode_engine"
    assert transcode_agent_service.TRANSCODE_MODULE_NAME == "fangzheng_web_app.transcode_agent_engine"

    standard_engine = transcode_service.load_transcode_module()
    agent_engine = transcode_agent_service.load_transcode_module()

    assert standard_engine is not agent_engine
    assert standard_engine.__name__ == "fangzheng_web_app.transcode_engine"
    assert agent_engine.__name__ == "fangzheng_web_app.transcode_agent_engine"
    assert not hasattr(standard_engine, "get_customer_size_override")
    assert hasattr(agent_engine, "get_customer_size_override")

    print("agent engine isolation smoke passed")


if __name__ == "__main__":
    main()
