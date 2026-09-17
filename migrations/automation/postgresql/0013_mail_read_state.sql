-- IMAP \Seen state mirrored locally.  The source of truth is the mailbox;
-- this column is refreshed during each mail synchronization.
ALTER TABLE mail_messages
    ADD COLUMN IF NOT EXISTS is_seen INTEGER NOT NULL DEFAULT 0;

ALTER TABLE mail_messages
    ADD COLUMN IF NOT EXISTS seen_updated_at TEXT NOT NULL DEFAULT '';

CREATE INDEX IF NOT EXISTS idx_mail_messages_account_seen
    ON mail_messages(account_id, is_seen);
