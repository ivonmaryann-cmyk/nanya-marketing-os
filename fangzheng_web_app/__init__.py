from flask import Flask

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - optional local convenience dependency
    load_dotenv = None

from .db import init_db
from .paths import ensure_storage_dirs
from .routes import bp
from .rules import ensure_default_rule_version
from .bomin_rules import ensure_default_bomin_rule_version
from .hushi_rules import ensure_default_hushi_rule_version
from .job_control import reconcile_interrupted_jobs
from .price_calculation_rules import ensure_default_price_rule_versions
from .shennan_rules import ensure_default_shennan_rule_version
from .transcode_rules import ensure_default_transcode_rule_version
from .transcode_agent_rules import ensure_default_transcode_agent_rule_version
from .transcode_semantic_rules import ensure_default_transcode_semantic_rule_version


def create_app() -> Flask:
    if load_dotenv is not None:
        load_dotenv()

    app = Flask(__name__, template_folder="../templates", static_folder="../static")
    app.config["SECRET_KEY"] = "fangzheng-web-app-dev-secret"
    # 库存明细一次上传两份库存表，当前业务样例合计约 43 MB；预留 multipart 开销。
    app.config["MAX_CONTENT_LENGTH"] = 120 * 1024 * 1024

    ensure_storage_dirs()
    init_db()
    reconcile_interrupted_jobs()
    ensure_default_rule_version()
    ensure_default_transcode_rule_version()
    ensure_default_transcode_agent_rule_version()
    ensure_default_transcode_semantic_rule_version()
    ensure_default_shennan_rule_version()
    ensure_default_hushi_rule_version()
    ensure_default_bomin_rule_version()
    ensure_default_price_rule_versions()

    app.register_blueprint(bp)
    return app
