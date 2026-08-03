from __future__ import annotations

import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app import db
from fangzheng_web_app.transcode_model_config import (
    load_user_model_config,
    update_user_model_config,
)


def main() -> None:
    original_path = db.DATABASE_PATH
    with tempfile.TemporaryDirectory() as tmp_dir:
        db.DATABASE_PATH = Path(tmp_dir) / "app.db"
        try:
            db.init_db()
            first = update_user_model_config(
                "user-a",
                enabled=True,
                base_url="https://api.deepseek.example",
                api_key="key-a",
                model="deepseek-v4-pro",
            )
            second = update_user_model_config(
                "user-b",
                enabled=False,
                base_url="https://api.deepseek.example",
                api_key="key-b",
                model="deepseek-chat",
            )
            assert first.enabled and first.api_key == "key-a"
            assert not second.enabled and second.api_key == "key-b"
            assert load_user_model_config("user-a").model == "deepseek-v4-pro"
            assert load_user_model_config("user-b").model == "deepseek-chat"
            preserved = update_user_model_config(
                "user-a",
                enabled=True,
                base_url="https://api.deepseek.example/v1",
                api_key=None,
                model="deepseek-v4-pro",
            )
            assert preserved.api_key == "key-a"
            assert preserved.base_url == "https://api.deepseek.example/v1"
        finally:
            db.DATABASE_PATH = original_path
    print("user model config smoke passed isolation=2 key-preserved=true")


if __name__ == "__main__":
    main()
