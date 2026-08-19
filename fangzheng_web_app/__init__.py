from flask import Flask

from .local_env import load_local_env
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
from .transcode_rule_center import ensure_daily_backup, ensure_rule_center_tables
from .pp_transcode_rules import ensure_pp_transcode_daily_backup, ensure_pp_transcode_tables, seed_pp_transcode_rules
from .mail_transcode_agent import bp as mail_transcode_bp


def create_app() -> Flask:
    load_local_env()
    app = Flask(__name__, template_folder="../templates", static_folder="../static")
    app.config["SECRET_KEY"] = "fangzheng-web-app-dev-secret"
    # 库存明细一次上传两份库存表，当前业务样例合计约 43 MB；预留 multipart 开销。
    app.config["MAX_CONTENT_LENGTH"] = 120 * 1024 * 1024

    ensure_storage_dirs()
    init_db()
    ensure_rule_center_tables()
    ensure_pp_transcode_tables()
    reconcile_interrupted_jobs()
    ensure_default_rule_version()
    ensure_default_transcode_rule_version()
    ensure_default_transcode_agent_rule_version()
    ensure_default_transcode_semantic_rule_version()
    ensure_daily_backup()
    seed_pp_transcode_rules()
    ensure_pp_transcode_daily_backup()
    ensure_default_shennan_rule_version()
    ensure_default_hushi_rule_version()
    ensure_default_bomin_rule_version()
    ensure_default_price_rule_versions()

    app.register_blueprint(bp)
    app.register_blueprint(mail_transcode_bp, url_prefix="/mail-transcode")
    return app
