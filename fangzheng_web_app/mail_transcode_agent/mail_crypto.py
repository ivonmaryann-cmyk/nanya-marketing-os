from __future__ import annotations

import base64
import hashlib
import os

from cryptography.fernet import Fernet, InvalidToken


MASTER_KEY_ENV = "MAIL_MASTER_KEY"
FALLBACK_MASTER_KEY_ENV = "PDF_EXCEL_AI_CONFIG_MASTER_KEY"


def _fernet() -> Fernet:
    raw_key = (
        os.environ.get(MASTER_KEY_ENV, "").strip()
        or os.environ.get(FALLBACK_MASTER_KEY_ENV, "").strip()
    )
    if raw_key:
        try:
            return Fernet(raw_key.encode("ascii"))
        except Exception:
            pass
    digest = hashlib.sha256(b"fangzheng-mail-transcode-agent-v1").digest()
    return Fernet(base64.urlsafe_b64encode(digest))


def encrypt_text(value: str) -> str:
    if not value:
        return ""
    return _fernet().encrypt(value.encode("utf-8")).decode("ascii")


def decrypt_text(ciphertext: str) -> str:
    if not ciphertext:
        return ""
    try:
        return _fernet().decrypt(ciphertext.encode("ascii")).decode("utf-8")
    except InvalidToken:
        return ""
