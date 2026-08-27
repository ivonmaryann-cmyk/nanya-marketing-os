from __future__ import annotations

import html
import re
import zipfile
from dataclasses import dataclass
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

from zhconv import convert


ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}
MAX_ARCHIVE_ENTRIES = 10_000
MAX_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_TEXT_LENGTH = 100_000

_PREFIX = r"(?:[A-Za-z_][\w.-]*:)?"
_TEXT_NODE_RE = re.compile(
    rf"(<{_PREFIX}t\b[^>]*>)(.*?)(</{_PREFIX}t\s*>)",
    flags=re.DOTALL,
)
_SHARED_ITEM_RE = re.compile(
    rf"(<{_PREFIX}si\b[^>]*>)(.*?)(</{_PREFIX}si\s*>)",
    flags=re.DOTALL,
)
_CELL_RE = re.compile(
    rf"(<{_PREFIX}c\b(?P<attrs>[^>]*)>)(?P<body>.*?)(</{_PREFIX}c\s*>)",
    flags=re.DOTALL,
)
_VALUE_RE = re.compile(rf"<{_PREFIX}v\b[^>]*>\s*(\d+)\s*</{_PREFIX}v\s*>")
_TYPE_RE = re.compile(r"\bt\s*=\s*['\"]([^'\"]+)['\"]")


class WorkbookConversionError(ValueError):
    pass


@dataclass(frozen=True)
class ConversionStats:
    sheet_count: int
    text_cell_count: int
    changed_cell_count: int
    changed_character_count: int

    def as_dict(self) -> dict[str, int]:
        return {
            "sheet_count": self.sheet_count,
            "text_cell_count": self.text_cell_count,
            "changed_cell_count": self.changed_cell_count,
            "changed_character_count": self.changed_character_count,
        }


@dataclass(frozen=True)
class WorkbookConversionResult:
    content: BytesIO
    filename: str
    mimetype: str
    stats: ConversionStats


def changed_character_count(before: str, after: str) -> int:
    if before == after:
        return 0
    matcher = SequenceMatcher(a=before, b=after, autojunk=False)
    return sum(
        max(source_end - source_start, target_end - target_start)
        for tag, source_start, source_end, target_start, target_end in matcher.get_opcodes()
        if tag != "equal"
    )


def convert_text_to_simplified(value: str) -> tuple[str, int]:
    if not isinstance(value, str):
        raise TypeError("value must be a string")
    converted = convert(value, "zh-cn").replace("\u00a0", " ")
    return converted, changed_character_count(value, converted)


def _convert_text_nodes(xml: str) -> tuple[str, int]:
    changed = 0

    def replace(match: re.Match[str]) -> str:
        nonlocal changed
        original = html.unescape(match.group(2))
        converted, count = convert_text_to_simplified(original)
        changed += count
        if not count:
            return match.group(0)
        return f"{match.group(1)}{escape(converted)}{match.group(3)}"

    return _TEXT_NODE_RE.sub(replace, xml), changed


def _convert_shared_strings(xml_bytes: bytes) -> tuple[bytes, dict[int, int]]:
    try:
        xml = xml_bytes.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise WorkbookConversionError("Excel 共享字符串不是有效的 UTF-8 内容") from exc
    changed_items: dict[int, int] = {}
    item_index = 0

    def replace(match: re.Match[str]) -> str:
        nonlocal item_index
        converted_body, changed = _convert_text_nodes(match.group(2))
        if changed:
            changed_items[item_index] = changed
        item_index += 1
        return f"{match.group(1)}{converted_body}{match.group(3)}"

    converted = _SHARED_ITEM_RE.sub(replace, xml)
    return converted.encode("utf-8"), changed_items


def _convert_worksheet(
    xml_bytes: bytes,
    shared_changes: dict[int, int],
) -> tuple[bytes, int, int, int]:
    try:
        xml = xml_bytes.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise WorkbookConversionError("Excel 工作表不是有效的 UTF-8 内容") from exc
    text_cells = 0
    changed_cells = 0
    changed_characters = 0

    def replace(match: re.Match[str]) -> str:
        nonlocal text_cells, changed_cells, changed_characters
        attrs = match.group("attrs")
        body = match.group("body")
        type_match = _TYPE_RE.search(attrs)
        cell_type = type_match.group(1) if type_match else ""
        if cell_type == "s":
            value_match = _VALUE_RE.search(body)
            if not value_match:
                return match.group(0)
            text_cells += 1
            shared_index = int(value_match.group(1))
            changed = shared_changes.get(shared_index, 0)
            if changed:
                changed_cells += 1
                changed_characters += changed
            return match.group(0)
        if cell_type != "inlineStr":
            return match.group(0)
        text_cells += 1
        converted_body, changed = _convert_text_nodes(body)
        if not changed:
            return match.group(0)
        changed_cells += 1
        changed_characters += changed
        return f"{match.group(1)}{converted_body}{match.group(4)}"

    converted = _CELL_RE.sub(replace, xml)
    return converted.encode("utf-8"), text_cells, changed_cells, changed_characters


def _safe_output_filename(filename: str, suffix: str) -> str:
    name = Path(str(filename or "Excel")).name
    stem = Path(name).stem.strip() or "Excel"
    stem = re.sub(r"[\x00-\x1f<>:\"/\\|?*]+", "_", stem).rstrip(" .") or "Excel"
    return f"{stem}_简体版{suffix}"


def _validate_archive(source: zipfile.ZipFile) -> list[zipfile.ZipInfo]:
    entries = source.infolist()
    if len(entries) > MAX_ARCHIVE_ENTRIES:
        raise WorkbookConversionError("Excel 内部文件数量异常，已停止处理")
    if len({entry.filename for entry in entries}) != len(entries):
        raise WorkbookConversionError("Excel 内部包含重复文件，无法安全处理")
    if any(entry.flag_bits & 0x1 for entry in entries):
        raise WorkbookConversionError("暂不支持加密的 Excel 文件")
    if sum(entry.file_size for entry in entries) > MAX_UNCOMPRESSED_BYTES:
        raise WorkbookConversionError("Excel 解压后体积过大，已停止处理")
    names = {entry.filename for entry in entries}
    if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
        raise WorkbookConversionError("文件不是有效的 Excel OOXML 工作簿")
    return entries


def convert_workbook_to_simplified(content: bytes, filename: str) -> WorkbookConversionResult:
    suffix = Path(str(filename or "")).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise WorkbookConversionError("仅支持 .xlsx 和 .xlsm 文件")
    if not content:
        raise WorkbookConversionError("上传的 Excel 文件为空")

    source_buffer = BytesIO(content)
    output_buffer = BytesIO()
    try:
        with zipfile.ZipFile(source_buffer, "r") as source:
            entries = _validate_archive(source)
            shared_changes: dict[int, int] = {}
            shared_name = "xl/sharedStrings.xml"
            converted_entries: dict[str, bytes] = {}
            if shared_name in source.namelist():
                converted_shared, shared_changes = _convert_shared_strings(source.read(shared_name))
                converted_entries[shared_name] = converted_shared

            sheet_count = 0
            text_cells = 0
            changed_cells = 0
            changed_characters = 0
            for entry in entries:
                if not re.fullmatch(r"xl/worksheets/[^/]+\.xml", entry.filename):
                    continue
                sheet_count += 1
                converted, texts, cells, characters = _convert_worksheet(
                    source.read(entry.filename), shared_changes
                )
                converted_entries[entry.filename] = converted
                text_cells += texts
                changed_cells += cells
                changed_characters += characters

            with zipfile.ZipFile(output_buffer, "w") as target:
                target.comment = source.comment
                for entry in entries:
                    data = converted_entries.get(entry.filename)
                    if data is None:
                        data = source.read(entry.filename)
                    target.writestr(entry, data)
    except zipfile.BadZipFile as exc:
        raise WorkbookConversionError("文件损坏或不是有效的 Excel OOXML 文件") from exc

    output_buffer.seek(0)
    mimetype = (
        "application/vnd.ms-excel.sheet.macroEnabled.12"
        if suffix == ".xlsm"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return WorkbookConversionResult(
        content=output_buffer,
        filename=_safe_output_filename(filename, suffix),
        mimetype=mimetype,
        stats=ConversionStats(
            sheet_count=sheet_count,
            text_cell_count=text_cells,
            changed_cell_count=changed_cells,
            changed_character_count=changed_characters,
        ),
    )
