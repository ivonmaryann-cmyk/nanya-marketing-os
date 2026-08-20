from __future__ import annotations

import sqlite3
from contextlib import closing, contextmanager
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterator


@contextmanager
def sqlite_snapshot(source: Path) -> Iterator[Path]:
    """Create a consistent read snapshot without changing or locking the source for migration."""
    if not source.is_file():
        raise FileNotFoundError(source)
    with TemporaryDirectory(prefix="nanya-automation-snapshot-") as directory:
        target = Path(directory) / "automation.db"
        source_uri = f"file:{source.resolve().as_posix()}?mode=ro"
        with closing(sqlite3.connect(source_uri, uri=True)) as source_connection:
            with closing(sqlite3.connect(target)) as target_connection:
                source_connection.backup(target_connection)
        yield target
