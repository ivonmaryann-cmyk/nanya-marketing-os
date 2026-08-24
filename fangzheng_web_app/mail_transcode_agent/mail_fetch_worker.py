from __future__ import annotations

"""Entry point for a queued mailbox fetch.

It intentionally has no Flask request context.  Both manual sync and a later
cron invocation use the same function, which keeps scheduling operationally
configurable without duplicating IMAP or persistence logic.
"""

import argparse

from .mail_fetch_service import fetch_latest_order_mails


def main() -> None:
    parser = argparse.ArgumentParser(description="Run one queued mailbox fetch")
    parser.add_argument("--task-id", required=True, type=int)
    parser.add_argument("--account-id", required=True, type=int)
    parser.add_argument("--owner-employee-id", required=True)
    parser.add_argument("--created-by", default="")
    parser.add_argument("--lookback-days", default=2, type=int)
    parser.add_argument("--limit", type=int)
    args = parser.parse_args()
    fetch_latest_order_mails(
        args.account_id,
        args.limit,
        created_by=args.created_by,
        owner_employee_id=args.owner_employee_id,
        lookback_days=args.lookback_days,
        fetch_task_id=args.task_id,
    )


if __name__ == "__main__":
    main()
