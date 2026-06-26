from flask import Flask

from .db import init_db
from .paths import ensure_storage_dirs
from .routes import bp
from .rules import ensure_default_rule_version
from .bomin_rules import ensure_default_bomin_rule_version
from .hushi_rules import ensure_default_hushi_rule_version
from .job_control import reconcile_interrupted_jobs
from .shennan_rules import ensure_default_shennan_rule_version
from .transcode_rules import ensure_default_transcode_rule_version


def create_app() -> Flask:
    app = Flask(__name__, template_folder="../templates", static_folder="../static")
    app.config["SECRET_KEY"] = "fangzheng-web-app-dev-secret"
    app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

    ensure_storage_dirs()
    init_db()
    reconcile_interrupted_jobs()
    ensure_default_rule_version()
    ensure_default_transcode_rule_version()
    ensure_default_shennan_rule_version()
    ensure_default_hushi_rule_version()
    ensure_default_bomin_rule_version()

    app.register_blueprint(bp)
    return app
