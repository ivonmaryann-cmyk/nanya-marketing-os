from __future__ import annotations

"""Queue one daily fetch per enabled business mailbox.

Install this module in the server's cron (for example every workday at 08:00).
It only queues workers and exits, so cron never holds an IMAP connection.
"""

from . import mail_store
from .mail_fetch_service import queue_latest_order_mails


def queue_enabled_business_mailboxes(*, lookback_days: int = 2) -> list[int]:
    task_ids: list[int] = []
    for account in mail_store.list_accounts():
        if not account.get("enabled"):
            continue
        owner = str(account.get("owner_employee_id") or "")
        if not owner:
            continue
        try:
            task = queue_latest_order_mails(
                int(account["id"]), created_by="scheduler", owner_employee_id=owner,
                lookback_days=lookback_days,
            )
        except ValueError:
            # An existing queued/running task is expected when cron overlaps a
            # manually requested sync; leave it intact rather than duplicating.
            continue
        task_ids.append(int(task["fetch_task_id"]))
    return task_ids


if __name__ == "__main__":
    queue_enabled_business_mailboxes()
