from __future__ import annotations

import os
import tempfile
import unittest
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import patch

from fangzheng_web_app.file_storage import (
    LocalFileStorage,
    automation_object_key,
    is_allowed_automation_path,
    resolve_attachment_path,
    save_automation_file,
)
from fangzheng_web_app.file_storage.manifest import _entry
from fangzheng_web_app.mail_transcode_agent.mail_fetch_service import _collect_attachments


class FailingStorage(LocalFileStorage):
    def save(self, stream, object_key, metadata=None):
        raise OSError("injected managed storage failure")


class FileStorageDualWriteTests(unittest.TestCase):
    def test_dual_write_is_atomic_and_read_prefers_managed_copy(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legacy_root = root / "legacy"
            managed_root = root / "managed"
            legacy_path = legacy_root / "mail_transcode" / "1" / "2" / "attachments" / "a.txt"
            result = save_automation_file(
                b"content",
                legacy_path,
                legacy_root=legacy_root,
                managed_storage=LocalFileStorage(managed_root),
                dual_write=True,
            )
            self.assertEqual("mail_transcode/1/2/attachments/a.txt", result.object_key)
            self.assertEqual(b"content", legacy_path.read_bytes())
            managed_path = managed_root / result.object_key
            self.assertEqual(b"content", managed_path.read_bytes())
            legacy_path.write_bytes(b"legacy")
            with patch.dict(os.environ, {"AUTOMATION_FILE_DUAL_READ_ENABLED": "true"}, clear=False):
                resolved = resolve_attachment_path(
                    str(legacy_path), legacy_root=legacy_root, managed_root=managed_root
                )
            self.assertEqual(managed_path.resolve(), resolved)

    def test_managed_failure_does_not_create_legacy_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legacy_root = root / "legacy"
            legacy_path = legacy_root / "mail_transcode" / "1" / "2" / "original.eml"
            with self.assertRaisesRegex(OSError, "injected"):
                save_automation_file(
                    b"mail",
                    legacy_path,
                    legacy_root=legacy_root,
                    managed_storage=FailingStorage(root / "managed"),
                    dual_write=True,
                )
            self.assertFalse(legacy_path.exists())

    def test_attachment_metadata_is_not_returned_when_storage_fails(self) -> None:
        message = EmailMessage()
        message.set_content("body")
        message.add_attachment(b"payload", maintype="application", subtype="pdf", filename="order.pdf")
        with tempfile.TemporaryDirectory() as directory, patch(
            "fangzheng_web_app.mail_transcode_agent.mail_fetch_service.save_automation_file",
            side_effect=OSError("injected managed storage failure"),
        ):
            with self.assertRaisesRegex(OSError, "injected"):
                _collect_attachments(message, Path(directory))

    def test_object_keys_and_download_roots_are_restricted(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legacy_root = root / "legacy"
            valid = legacy_root / "mail_transcode" / "1" / "original.eml"
            self.assertEqual(
                "mail_transcode/1/original.eml",
                automation_object_key(valid, legacy_root=legacy_root),
            )
            with self.assertRaisesRegex(ValueError, "outside"):
                automation_object_key(root / "outside.txt", legacy_root=legacy_root)
            managed_root = root / "managed"
            managed_file = managed_root / "mail_transcode" / "1" / "file.txt"
            managed_file.parent.mkdir(parents=True)
            managed_file.write_text("ok", encoding="utf-8")
            with patch.dict(os.environ, {"AUTOMATION_FILE_STORAGE_ROOT": str(managed_root)}, clear=False):
                self.assertTrue(is_allowed_automation_path(managed_file))
                self.assertFalse(is_allowed_automation_path(root / "outside.txt"))

    def test_inventory_distinguishes_pending_verified_and_failed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legacy_root = root / "legacy"
            managed = LocalFileStorage(root / "managed")
            source = legacy_root / "mail_transcode" / "1" / "2" / "original.eml"
            source.parent.mkdir(parents=True)
            source.write_bytes(b"mail")
            pending = _entry(
                kind="original_eml", record_id=1, legacy_path=str(source),
                managed=managed, legacy_root=legacy_root,
            )
            self.assertEqual("pending", pending["status"])
            with source.open("rb") as stream:
                managed.save(stream, pending["object_key"])
            verified = _entry(
                kind="original_eml", record_id=1, legacy_path=str(source),
                managed=managed, legacy_root=legacy_root,
            )
            self.assertEqual("verified", verified["status"])
            failed = _entry(
                kind="attachment", record_id=2, legacy_path=str(source), expected_size=99,
                managed=managed, legacy_root=legacy_root,
            )
            self.assertEqual("failed", failed["status"])


if __name__ == "__main__":
    unittest.main()
