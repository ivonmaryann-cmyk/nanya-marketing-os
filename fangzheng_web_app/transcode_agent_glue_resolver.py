from __future__ import annotations

import re
import threading
import unicodedata
from collections import OrderedDict
from dataclasses import dataclass
from typing import Any

from .transcode_customer_identity import customer_names_match


VALID_GLUE_CODE = re.compile(r"^[A-Z0-9]{2}$")
_INDEX_CACHE_LIMIT = 16
_INDEX_CACHE: OrderedDict[int, tuple[dict[str, list[dict]], "_GlueIndex"]] = OrderedDict()
_INDEX_CACHE_LOCK = threading.RLock()


@dataclass(frozen=True)
class _NameEntry:
    row: dict
    pattern: re.Pattern[str]
    length: int


@dataclass(frozen=True)
class _SelectionEntry:
    row: dict
    name_pattern: re.Pattern[str]
    keyword_patterns: tuple[re.Pattern[str], ...]
    priority: int


@dataclass(frozen=True)
class _GlueIndex:
    id_pattern: re.Pattern[str] | None
    id_rows: dict[str, tuple[dict, ...]]
    master_names: tuple[_NameEntry, ...]
    aliases: tuple[_NameEntry, ...]
    legacy_names: frozenset[str]
    selections: tuple[_SelectionEntry, ...]


def resolve_agent_glue(
    mapping_tables: dict[str, list[dict]],
    *,
    spec: str,
    context: str = "",
    customer_code: str = "",
    customer_name: str = "",
    current_code: str = "",
) -> dict[str, Any] | None:
    source = f"{spec or ''} {context or ''}".strip()
    normalized_source = unicodedata.normalize("NFKC", source).upper()
    current = str(current_code or "").strip().upper()
    current_valid = bool(VALID_GLUE_CODE.fullmatch(current))
    index = _get_glue_index(mapping_tables)

    explicit_ids = _find_explicit_ids(index, source)
    if explicit_ids:
        return _resolve_rows(explicit_ids, "新版胶系编号精确命中")

    selection = _match_selection_rule(
        index.selections,
        normalized_source=normalized_source,
        customer_code=customer_code,
        customer_name=customer_name,
    )
    if selection and _has_runtime_selection_condition(selection):
        return {
            "status": "matched",
            "code": str(selection.get("输出胶系代码") or "").strip().upper(),
            "rule_id": selection.get("映射ID", ""),
            "source": "Agent胶系选择规则",
            "source_row": "",
            "name": selection.get("胶系名称", ""),
            "text": selection.get("规则文本", ""),
        }

    matched_master = _longest_name_matches(index.master_names, normalized_source)
    if matched_master:
        codes = {str(row.get("输出胶系代码") or "").strip().upper() for row in matched_master}
        if len(codes) > 1:
            preferred = min(
                matched_master,
                key=lambda row: _source_row_number(row.get("来源行号")),
            )
            candidates = [
                {
                    "code": str(row.get("输出胶系代码") or "").strip().upper(),
                    "glue_id": str(row.get("胶系编号") or "").strip(),
                    "source_row": str(row.get("来源行号") or "").strip(),
                }
                for row in sorted(
                    matched_master,
                    key=lambda row: _source_row_number(row.get("来源行号")),
                )
            ]
            result = _resolve_rows([preferred], "最新版胶系名称优先命中")
            result["uncertain"] = True
            result["candidates"] = candidates
            result["conflict"] = (
                f"最新版胶系主表同名多码待补口径："
                f"{preferred.get('胶系名称', '')}候选{'/'.join(sorted(codes))}；"
                f"当前按来源行{preferred.get('来源行号', '')}优先使用"
                f"{preferred.get('输出胶系代码', '')}"
            )
            return result
        return _resolve_rows(matched_master, "最新版胶系名称精确命中")

    if not current_valid:
        matched_aliases = _longest_name_matches(index.aliases, normalized_source)
        if matched_aliases:
            row = matched_aliases[0]
            return {
                "status": "matched",
                "code": str(row.get("输出胶系代码") or "").strip().upper(),
                "rule_id": row.get("映射ID", ""),
                "source": "Agent胶系兼容别名",
                "source_row": "",
                "name": row.get("标准胶系名称", "") or row.get("兼容名称", ""),
                "text": row.get("规则文本", ""),
            }
    return None


def clear_agent_glue_index_cache() -> None:
    with _INDEX_CACHE_LOCK:
        _INDEX_CACHE.clear()


def _get_glue_index(mapping_tables: dict[str, list[dict]]) -> _GlueIndex:
    cache_key = id(mapping_tables)
    with _INDEX_CACHE_LOCK:
        cached = _INDEX_CACHE.get(cache_key)
        if cached and cached[0] is mapping_tables:
            _INDEX_CACHE.move_to_end(cache_key)
            return cached[1]
        index = _build_glue_index(mapping_tables)
        _INDEX_CACHE[cache_key] = (mapping_tables, index)
        _INDEX_CACHE.move_to_end(cache_key)
        while len(_INDEX_CACHE) > _INDEX_CACHE_LIMIT:
            _INDEX_CACHE.popitem(last=False)
        return index


def _build_glue_index(mapping_tables: dict[str, list[dict]]) -> _GlueIndex:
    master_rows = [
        row
        for row in mapping_tables.get("Agent胶系主表", [])
        if _enabled(row) and _valid_code(row.get("输出胶系代码"))
    ]
    id_rows: dict[str, list[dict]] = {}
    for row in master_rows:
        glue_id = str(row.get("胶系编号") or "").strip().upper()
        if glue_id:
            id_rows.setdefault(glue_id, []).append(row)
    id_pattern = None
    if id_rows:
        alternatives = "|".join(
            re.escape(glue_id) for glue_id in sorted(id_rows, key=len, reverse=True)
        )
        id_pattern = re.compile(
            rf"(?<![A-Z0-9])(?:{alternatives})(?![A-Z0-9])",
            re.IGNORECASE,
        )

    aliases = [
        row
        for row in mapping_tables.get("Agent胶系兼容别名", [])
        if _enabled(row) and _valid_code(row.get("输出胶系代码"))
    ]
    selections = []
    for row in mapping_tables.get("Agent胶系选择规则", []):
        if not _enabled(row) or not _valid_code(row.get("输出胶系代码")):
            continue
        name_pattern = _compile_phrase_pattern(row.get("胶系名称"))
        if not name_pattern:
            continue
        keyword_patterns = tuple(
            pattern
            for item in re.split(r"[,，、;；]+", str(row.get("条件关键词") or ""))
            if item.strip()
            for pattern in [_compile_phrase_pattern(item.strip())]
            if pattern
        )
        selections.append(
            _SelectionEntry(
                row=row,
                name_pattern=name_pattern,
                keyword_patterns=keyword_patterns,
                priority=_priority(row.get("优先级")),
            )
        )
    selections.sort(key=lambda entry: entry.priority, reverse=True)
    return _GlueIndex(
        id_pattern=id_pattern,
        id_rows={key: tuple(value) for key, value in id_rows.items()},
        master_names=_compile_name_entries(master_rows, "胶系名称"),
        aliases=_compile_name_entries(aliases, "兼容名称"),
        legacy_names=frozenset(
            _normalize(row.get("兼容名称"))
            for row in aliases
            if _normalize(row.get("兼容名称"))
        ),
        selections=tuple(selections),
    )


def _compile_name_entries(rows: list[dict], field: str) -> tuple[_NameEntry, ...]:
    entries = []
    for row in rows:
        pattern = _compile_phrase_pattern(row.get(field))
        length = len(_normalize(row.get(field)))
        if pattern and length:
            entries.append(_NameEntry(row=row, pattern=pattern, length=length))
    entries.sort(key=lambda entry: entry.length, reverse=True)
    return tuple(entries)


def _find_explicit_ids(index: _GlueIndex, source: str) -> list[dict]:
    if not index.id_pattern:
        return []
    matched_rows = []
    seen_ids = set()
    for match in index.id_pattern.finditer(source):
        glue_id = match.group(0).upper()
        if glue_id in seen_ids:
            continue
        seen_ids.add(glue_id)
        matched_rows.extend(index.id_rows.get(glue_id, ()))
    return matched_rows


def _resolve_rows(rows: list[dict], source: str) -> dict[str, Any]:
    codes = {str(row.get("输出胶系代码") or "").strip().upper() for row in rows}
    if len(codes) != 1:
        return {
            "status": "conflict",
            "conflict": f"同一胶系命中多个代码：{'/'.join(sorted(codes))}",
        }
    row = rows[0]
    return {
        "status": "matched",
        "code": next(iter(codes)),
        "rule_id": row.get("映射ID", ""),
        "source": source,
        "source_row": row.get("来源行号", ""),
        "name": row.get("胶系名称", ""),
        "classification": row.get("胶系分类", ""),
        "glue_id": row.get("胶系编号", ""),
        "text": (
            f"{row.get('胶系编号', '')} {row.get('胶系名称', '')}"
            f" → {row.get('输出胶系代码', '')}"
        ).strip(),
    }


def _match_selection_rule(
    entries: tuple[_SelectionEntry, ...],
    *,
    normalized_source: str,
    customer_code: str,
    customer_name: str,
) -> dict | None:
    candidates: list[_SelectionEntry] = []
    for entry in entries:
        row = entry.row
        if not entry.name_pattern.search(normalized_source):
            continue
        if not _selection_customer_matches(row, customer_code, customer_name):
            continue
        if entry.keyword_patterns and not any(
            pattern.search(normalized_source) for pattern in entry.keyword_patterns
        ):
            continue
        candidates.append(entry)
    if not candidates:
        return None
    top_priority = candidates[0].priority
    top = [entry for entry in candidates if entry.priority == top_priority]
    if len({str(entry.row.get("输出胶系代码") or "").upper() for entry in top}) > 1:
        return None
    return top[0].row


def _selection_customer_matches(row: dict, customer_code: str, customer_name: str) -> bool:
    configured_code = str(row.get("条件客户代码") or "").strip()
    configured_name = str(row.get("条件客户简称") or "").strip()
    if configured_code:
        allowed = set(re.findall(r"\d+", configured_code))
        actual = set(re.findall(r"\d+", str(customer_code or "")))
        if not actual or allowed.isdisjoint(actual):
            return False
    if configured_name and not customer_names_match(configured_name, customer_name):
        return False
    return True


def _has_runtime_selection_condition(row: dict) -> bool:
    return any(
        str(row.get(field) or "").strip()
        for field in ("条件客户代码", "条件客户简称", "条件关键词")
    )


def _longest_name_matches(
    entries: tuple[_NameEntry, ...], normalized_source: str
) -> list[dict]:
    matches = []
    longest = 0
    for entry in entries:
        if longest and entry.length < longest:
            break
        if entry.pattern.search(normalized_source):
            longest = entry.length
            matches.append(entry.row)
    return matches


def _phrase_present(source: str, value: Any) -> bool:
    pattern = _compile_phrase_pattern(value)
    if not pattern:
        return False
    return bool(pattern.search(unicodedata.normalize("NFKC", str(source or "")).upper()))


def _compile_phrase_pattern(value: Any) -> re.Pattern[str] | None:
    raw = unicodedata.normalize("NFKC", str(value or "")).upper().strip()
    if not raw:
        return None
    characters = [re.escape(char) for char in raw if not char.isspace() and char not in "_-"]
    if not characters:
        return None
    pattern = r"[\s_\-]*".join(characters)
    if raw[-1].isalnum():
        pattern += r"(?![A-Z0-9])"
    return re.compile(pattern)


def _token_present(source: str, value: Any) -> bool:
    token = str(value or "").strip()
    if not token:
        return False
    return bool(re.search(rf"(?<![A-Z0-9]){re.escape(token)}(?![A-Z0-9])", source, re.IGNORECASE))


def _enabled(row: dict) -> bool:
    return str(row.get("启用") or "").strip().upper() in {"是", "Y", "YES", "TRUE", "1"}


def _valid_code(value: Any) -> bool:
    return bool(VALID_GLUE_CODE.fullmatch(str(value or "").strip().upper()))


def _priority(value: Any) -> int:
    try:
        return int(float(str(value or "0")))
    except ValueError:
        return 0


def _source_row_number(value: Any) -> int:
    try:
        return int(float(str(value or "999999")))
    except ValueError:
        return 999999


def _normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).upper()
    text = text.replace("（", "(").replace("）", ")")
    return re.sub(r"[\s_\-]+", "", text)
