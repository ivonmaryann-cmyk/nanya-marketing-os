from __future__ import annotations

from pathlib import Path


def safe_unlink(path_value: str | None) -> bool:
    if not path_value:
        return True
    path = Path(path_value)
    if not path.exists():
        return True
    try:
        path.unlink(missing_ok=True)
        return True
    except PermissionError:
        return False
