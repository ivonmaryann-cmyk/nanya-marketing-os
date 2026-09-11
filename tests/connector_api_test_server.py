"""Minimal HTTP host used only by the Node MCP contract test.

It registers the real connector blueprint but deliberately does not call
``create_app``/database initialization, because initialize/tools-list require
no business data and must be testable independently of external PostgreSQL.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from flask import Flask

# Executing a file under tests/ makes Python put tests/ first on sys.path.
# Add the repository root explicitly so this helper also works from Node spawn.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from fangzheng_web_app.connector_api import bp

app = Flask(__name__)
app.register_blueprint(bp)

if __name__ == "__main__":
    port = int(os.getenv("CONNECTOR_TEST_PORT", "5011"))
    print(f"READY http://127.0.0.1:{port}", flush=True)
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)
