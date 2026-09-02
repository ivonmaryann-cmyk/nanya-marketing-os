-- Mailbox credentials remain private to their owner. Administrators may keep a
-- separate configuration for an email another user has already configured.
ALTER TABLE mail_accounts DROP CONSTRAINT IF EXISTS mail_accounts_email_key;
DROP INDEX IF EXISTS mail_accounts_email_key;

CREATE UNIQUE INDEX IF NOT EXISTS idx_mail_accounts_owner_email
    ON mail_accounts(owner_employee_id, email);
