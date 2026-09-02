-- SMTP reply configuration belongs to the mailbox that owns the source email.
-- Secrets remain encrypted by the application before they reach this column.
ALTER TABLE mail_accounts ADD COLUMN IF NOT EXISTS smtp_host TEXT NOT NULL DEFAULT '';
ALTER TABLE mail_accounts ADD COLUMN IF NOT EXISTS smtp_port INTEGER NOT NULL DEFAULT 465;
ALTER TABLE mail_accounts ADD COLUMN IF NOT EXISTS smtp_security TEXT NOT NULL DEFAULT 'ssl';
ALTER TABLE mail_accounts ADD COLUMN IF NOT EXISTS smtp_username TEXT NOT NULL DEFAULT '';
ALTER TABLE mail_accounts ADD COLUMN IF NOT EXISTS smtp_auth_code_ciphertext TEXT NOT NULL DEFAULT '';
ALTER TABLE mail_accounts ADD COLUMN IF NOT EXISTS smtp_sender_name TEXT NOT NULL DEFAULT '';
ALTER TABLE mail_accounts ADD COLUMN IF NOT EXISTS smtp_enabled INTEGER NOT NULL DEFAULT 0;
ALTER TABLE mail_accounts ADD COLUMN IF NOT EXISTS smtp_last_test_at TEXT;
ALTER TABLE mail_accounts ADD COLUMN IF NOT EXISTS smtp_last_test_status TEXT NOT NULL DEFAULT '';
