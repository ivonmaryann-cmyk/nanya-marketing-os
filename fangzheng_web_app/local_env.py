"""加载不进入版本库的本地服务配置。"""

from __future__ import annotations

import os
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOCAL_ENV_PATH = PROJECT_ROOT / "config" / "local.env"


def load_local_env(path: Path = LOCAL_ENV_PATH) -> None:
    """仅补充当前进程未设置的变量，显式环境变量始终优先。"""
    if not path.is_file():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if not key or key in os.environ:
            continue
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            value = value[1:-1]
        os.environ[key] = value
