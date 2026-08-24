from __future__ import annotations

"""Run a first-time order-entry extraction outside the web request process."""

import argparse

from .local_env import load_local_env
from .order_entry_service import run_template_extraction_task


def main() -> None:
    parser = argparse.ArgumentParser(description="Run one queued order-entry extraction")
    parser.add_argument("--task-id", required=True, type=int)
    parser.add_argument("--case-id", required=True, type=int)
    parser.add_argument("--employee-id", required=True)
    args = parser.parse_args()
    # A worker is started as an independent interpreter, so it must load the
    # local PostgreSQL configuration just as the Flask application does.
    load_local_env()
    run_template_extraction_task(args.task_id, args.case_id, args.employee_id)


if __name__ == "__main__":
    main()
