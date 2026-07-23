from __future__ import annotations

import hashlib
import json
import os
import sqlite3
from dataclasses import dataclass, field
from typing import Any, Mapping
from urllib.parse import urlparse

from cryptography.fernet import Fernet, InvalidToken

from .db import db_cursor, utcnow

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - optional local convenience dependency
    load_dotenv = None


if load_dotenv is not None:
    load_dotenv()


ACTIVE_VERSION_SETTING = "active_pdf_excel_ai_config_version"
MASTER_KEY_ENV = "PDF_EXCEL_AI_CONFIG_MASTER_KEY"
MAX_INSTRUCTION_LENGTH = 4000

DEFAULT_REPAIR_INSTRUCTION = "优先补足有原文证据的空字段；已有正确值不得覆盖。"
DEFAULT_REBUILD_INSTRUCTION = "只重建能够从订单正文逐行确认的物料明细。"
DEFAULT_HEADER_MAPPING_INSTRUCTION = "优先按客户明确表头语义映射，不根据列位置猜测。"


class AiConfigError(ValueError):
    pass


class AiConfigConflictError(AiConfigError):
    pass


class AiConfigEncryptionError(AiConfigError):
    pass


def _env_flag(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def _env_int(name: str, default: int, *, minimum: int, maximum: int) -> int:
    raw = os.environ.get(name, "").strip()
    try:
        value = int(raw)
    except (TypeError, ValueError):
        value = default
    return max(minimum, min(value, maximum))


def _enabled_value(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class AiRepairConfig:
    enabled: bool
    api_key: str = field(repr=False)
    base_url: str
    model: str
    timeout_seconds: int
    max_rows: int
    version_id: int | None = None
    source: str = "environment"
    repair_instruction: str = DEFAULT_REPAIR_INSTRUCTION
    rebuild_instruction: str = DEFAULT_REBUILD_INSTRUCTION
    header_mapping_instruction: str = DEFAULT_HEADER_MAPPING_INSTRUCTION
    load_error: str = ""

    @property
    def available(self) -> bool:
        return self.enabled and bool(self.api_key) and not self.load_error

    @property
    def prompt_digest(self) -> str:
        payload = "\n".join(
            [self.repair_instruction, self.rebuild_instruction, self.header_mapping_instruction]
        )
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]

    @property
    def fingerprint(self) -> str:
        payload = {
            "enabled": self.enabled,
            "base_url": self.base_url,
            "model": self.model,
            "timeout_seconds": self.timeout_seconds,
            "max_rows": self.max_rows,
            "key_digest": hashlib.sha256(self.api_key.encode("utf-8")).hexdigest() if self.api_key else "",
            "prompt_digest": self.prompt_digest,
        }
        encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return hashlib.sha256(encoded.encode("utf-8")).hexdigest()[:24]

    def safe_metadata(self) -> dict[str, Any]:
        return {
            "version_id": self.version_id,
            "source": self.source,
            "enabled": self.enabled,
            "available": self.available,
            "key_configured": bool(self.api_key),
            "base_url": self.base_url,
            "model": self.model,
            "timeout_seconds": self.timeout_seconds,
            "max_rows": self.max_rows,
            "fingerprint": self.fingerprint,
            "prompt_digest": self.prompt_digest,
            "load_error": self.load_error,
        }

    def safe_status(self) -> str:
        version = f"v{self.version_id}" if self.version_id is not None else "环境变量"
        if self.load_error:
            return f"AI补缺配置 {version} 无法读取，已自动跳过：{self.load_error}"
        if not self.enabled:
            return f"AI补缺未启用（配置：{version}）。"
        if not self.api_key:
            return f"AI补缺已启用（配置：{version}），但缺少 API Key，已自动跳过。"
        return (
            f"AI补缺已启用：配置={version}，model={self.model}，"
            f"timeout={self.timeout_seconds}s，max_rows={self.max_rows}。"
        )


def _legacy_env_config() -> AiRepairConfig:
    return AiRepairConfig(
        enabled=_env_flag("PDF_EXCEL_AI_REPAIR_ENABLED", default=False),
        api_key=os.environ.get("DEEPSEEK_API_KEY", "").strip(),
        base_url=os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip().rstrip("/"),
        model=os.environ.get("DEEPSEEK_MODEL", "deepseek-v4-flash").strip() or "deepseek-v4-flash",
        timeout_seconds=_env_int("PDF_EXCEL_AI_REPAIR_TIMEOUT_SECONDS", 45, minimum=5, maximum=300),
        max_rows=_env_int("PDF_EXCEL_AI_REPAIR_MAX_ROWS", 12, minimum=1, maximum=50),
    )


def master_key_configured() -> bool:
    try:
        _fernet()
    except AiConfigEncryptionError:
        return False
    return True


def _fernet() -> Fernet:
    raw_key = os.environ.get(MASTER_KEY_ENV, "").strip()
    if not raw_key:
        raise AiConfigEncryptionError(f"服务器未配置 {MASTER_KEY_ENV}。")
    try:
        return Fernet(raw_key.encode("ascii"))
    except Exception as exc:
        raise AiConfigEncryptionError(f"{MASTER_KEY_ENV} 格式无效。") from exc


def _encrypt_api_key(api_key: str) -> str:
    if not api_key:
        return ""
    return _fernet().encrypt(api_key.encode("utf-8")).decode("ascii")


def _decrypt_api_key(ciphertext: str) -> str:
    if not ciphertext:
        return ""
    try:
        return _fernet().decrypt(ciphertext.encode("ascii")).decode("utf-8")
    except InvalidToken as exc:
        raise AiConfigEncryptionError("API Key 无法使用当前主密钥解密。") from exc


def _active_version_id_from_value(value: Any) -> int | None:
    try:
        version_id = int(str(value or "").strip())
    except (TypeError, ValueError):
        return None
    return version_id if version_id > 0 else None


def get_active_ai_config_version_id() -> int | None:
    try:
        with db_cursor() as conn:
            row = conn.execute("SELECT value FROM settings WHERE key = ?", (ACTIVE_VERSION_SETTING,)).fetchone()
    except sqlite3.OperationalError:
        return None
    return _active_version_id_from_value(row["value"] if row else "")


def _config_from_row(row: Any, *, strict: bool = False) -> AiRepairConfig:
    load_error = ""
    try:
        api_key = _decrypt_api_key(str(row["api_key_ciphertext"] or ""))
    except AiConfigEncryptionError as exc:
        if strict:
            raise
        api_key = ""
        load_error = str(exc)
    return AiRepairConfig(
        enabled=bool(row["enabled"]),
        api_key=api_key,
        base_url=str(row["base_url"] or "").rstrip("/"),
        model=str(row["model"] or ""),
        timeout_seconds=int(row["timeout_seconds"]),
        max_rows=int(row["max_rows"]),
        version_id=int(row["id"]),
        source="database",
        repair_instruction=str(row["repair_instruction"] or ""),
        rebuild_instruction=str(row["rebuild_instruction"] or ""),
        header_mapping_instruction=str(row["header_mapping_instruction"] or ""),
        load_error=load_error,
    )


def get_ai_repair_config(version_id: int | None = None, *, strict: bool = False) -> AiRepairConfig:
    requested_version = version_id if version_id is not None else get_active_ai_config_version_id()
    if requested_version is None:
        return _legacy_env_config()
    try:
        with db_cursor() as conn:
            row = conn.execute(
                "SELECT * FROM pdf_excel_ai_config_versions WHERE id = ?",
                (requested_version,),
            ).fetchone()
    except sqlite3.OperationalError:
        if version_id is None:
            return _legacy_env_config()
        row = None
    if not row:
        message = f"AI 配置版本 {requested_version} 不存在。"
        if strict:
            raise AiConfigError(message)
        fallback = _legacy_env_config()
        return AiRepairConfig(
            enabled=False,
            api_key="",
            base_url=fallback.base_url,
            model=fallback.model,
            timeout_seconds=fallback.timeout_seconds,
            max_rows=fallback.max_rows,
            version_id=requested_version,
            source="database",
            load_error=message,
        )
    return _config_from_row(row, strict=strict)


def validate_ai_config_input(values: Mapping[str, Any], fallback: AiRepairConfig | None = None) -> AiRepairConfig:
    fallback = fallback or get_ai_repair_config()
    enabled = _enabled_value(values.get("enabled"))
    base_url = str(values.get("base_url") or fallback.base_url).strip().rstrip("/")
    parsed_url = urlparse(base_url)
    if parsed_url.scheme.lower() != "https" or not parsed_url.netloc:
        raise AiConfigError("接口地址必须是有效的 HTTPS 地址。")
    model = str(values.get("model") or "").strip()
    if not model or len(model) > 120:
        raise AiConfigError("模型名称不能为空且不能超过 120 个字符。")
    try:
        timeout_seconds = int(str(values.get("timeout_seconds") or "").strip())
    except ValueError as exc:
        raise AiConfigError("超时时间必须是整数。") from exc
    if not 5 <= timeout_seconds <= 300:
        raise AiConfigError("超时时间必须在 5 到 300 秒之间。")
    try:
        max_rows = int(str(values.get("max_rows") or "").strip())
    except ValueError as exc:
        raise AiConfigError("最大候选行数必须是整数。") from exc
    if not 1 <= max_rows <= 50:
        raise AiConfigError("最大候选行数必须在 1 到 50 之间。")

    instructions = {
        "repair_instruction": str(values.get("repair_instruction") if values.get("repair_instruction") is not None else fallback.repair_instruction).strip(),
        "rebuild_instruction": str(values.get("rebuild_instruction") if values.get("rebuild_instruction") is not None else fallback.rebuild_instruction).strip(),
        "header_mapping_instruction": str(values.get("header_mapping_instruction") if values.get("header_mapping_instruction") is not None else fallback.header_mapping_instruction).strip(),
    }
    for instruction in instructions.values():
        if len(instruction) > MAX_INSTRUCTION_LENGTH:
            raise AiConfigError(f"每段业务指令不能超过 {MAX_INSTRUCTION_LENGTH} 个字符。")

    submitted_key = str(values.get("api_key") or "").strip()
    api_key = submitted_key or fallback.api_key
    if enabled and not api_key:
        raise AiConfigError("启用 AI 时必须配置 API Key。")
    return AiRepairConfig(
        enabled=enabled,
        api_key=api_key,
        base_url=base_url,
        model=model,
        timeout_seconds=timeout_seconds,
        max_rows=max_rows,
        source="web",
        **instructions,
    )


def ai_config_form_values(config: AiRepairConfig) -> dict[str, Any]:
    return {
        "enabled": config.enabled,
        "base_url": config.base_url,
        "model": config.model,
        "timeout_seconds": config.timeout_seconds,
        "max_rows": config.max_rows,
        "repair_instruction": config.repair_instruction,
        "rebuild_instruction": config.rebuild_instruction,
        "header_mapping_instruction": config.header_mapping_instruction,
        "api_key": "",
    }


def save_ai_config_version(
    config: AiRepairConfig,
    *,
    employee_id: str,
    expected_active_version_id: int | None,
    test_status: str,
    test_message: str,
    source_version_id: int | None = None,
) -> AiRepairConfig:
    if config.enabled and test_status != "passed":
        raise AiConfigError("启用 AI 的配置必须先通过连接测试。")
    ciphertext = _encrypt_api_key(config.api_key)
    now = utcnow()
    tested_at = now if test_status == "passed" else None
    with db_cursor() as conn:
        conn.execute("BEGIN IMMEDIATE")
        active_row = conn.execute(
            "SELECT value FROM settings WHERE key = ?", (ACTIVE_VERSION_SETTING,)
        ).fetchone()
        current_active = _active_version_id_from_value(active_row["value"] if active_row else "")
        if current_active != expected_active_version_id:
            raise AiConfigConflictError("活动配置已被其他管理员更新，请刷新页面后重试。")
        cursor = conn.execute(
            """
            INSERT INTO pdf_excel_ai_config_versions (
                enabled, base_url, model, timeout_seconds, max_rows, api_key_ciphertext,
                repair_instruction, rebuild_instruction, header_mapping_instruction,
                config_fingerprint, prompt_digest, test_status, test_message, tested_at,
                created_by, created_at, activated_by, activated_at, source_version_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                1 if config.enabled else 0,
                config.base_url,
                config.model,
                config.timeout_seconds,
                config.max_rows,
                ciphertext,
                config.repair_instruction,
                config.rebuild_instruction,
                config.header_mapping_instruction,
                config.fingerprint,
                config.prompt_digest,
                test_status,
                str(test_message or "")[:500],
                tested_at,
                employee_id,
                now,
                employee_id,
                now,
                source_version_id,
            ),
        )
        version_id = int(cursor.lastrowid)
        conn.execute(
            """
            INSERT INTO settings(key, value) VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (ACTIVE_VERSION_SETTING, str(version_id)),
        )
    return get_ai_repair_config(version_id, strict=True)


def list_ai_config_versions(limit: int = 50) -> list[dict[str, Any]]:
    active_id = get_active_ai_config_version_id()
    try:
        with db_cursor() as conn:
            rows = conn.execute(
                "SELECT * FROM pdf_excel_ai_config_versions ORDER BY id DESC LIMIT ?",
                (max(1, min(int(limit), 200)),),
            ).fetchall()
    except sqlite3.OperationalError:
        return []
    return [
        {
            "id": int(row["id"]),
            "active": int(row["id"]) == active_id,
            "enabled": bool(row["enabled"]),
            "base_url": row["base_url"],
            "model": row["model"],
            "timeout_seconds": row["timeout_seconds"],
            "max_rows": row["max_rows"],
            "key_configured": bool(row["api_key_ciphertext"]),
            "fingerprint": row["config_fingerprint"],
            "prompt_digest": row["prompt_digest"],
            "test_status": row["test_status"],
            "test_message": row["test_message"],
            "tested_at": row["tested_at"],
            "created_by": row["created_by"],
            "created_at": row["created_at"],
            "activated_by": row["activated_by"],
            "activated_at": row["activated_at"],
            "source_version_id": row["source_version_id"],
        }
        for row in rows
    ]


def config_from_manifest_snapshot(snapshot: Mapping[str, Any] | None) -> AiRepairConfig:
    snapshot = snapshot or {}
    version_id = _active_version_id_from_value(snapshot.get("version_id"))
    if version_id is not None:
        return get_ai_repair_config(version_id)
    legacy = _legacy_env_config()
    if snapshot.get("source") != "environment":
        return legacy
    try:
        return AiRepairConfig(
            enabled=bool(snapshot.get("enabled")),
            api_key=legacy.api_key,
            base_url=str(snapshot.get("base_url") or legacy.base_url),
            model=str(snapshot.get("model") or legacy.model),
            timeout_seconds=int(snapshot.get("timeout_seconds") or legacy.timeout_seconds),
            max_rows=int(snapshot.get("max_rows") or legacy.max_rows),
            source="environment",
            repair_instruction=legacy.repair_instruction,
            rebuild_instruction=legacy.rebuild_instruction,
            header_mapping_instruction=legacy.header_mapping_instruction,
        )
    except (TypeError, ValueError):
        return legacy
