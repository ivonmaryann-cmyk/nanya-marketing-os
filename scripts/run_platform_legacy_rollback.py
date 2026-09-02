"""Start the application against the retained pre-cutover data sources.

Use only when the formal migration runbook calls for rollback.  It does not
edit config/local.env, the unified PostgreSQL target, or either source store.
"""

from fangzheng_web_app.platform_rollback import configure_legacy_rollback_environment


def main() -> None:
    configure_legacy_rollback_environment()
    from fangzheng_web_app import create_app

    create_app().run(host="0.0.0.0", port=5000, debug=True)


if __name__ == "__main__":
    main()
