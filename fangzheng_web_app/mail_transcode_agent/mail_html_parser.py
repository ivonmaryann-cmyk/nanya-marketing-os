from __future__ import annotations

import html as html_lib
import re
from html.parser import HTMLParser

from zhconv import convert


CUSTOMER_ALIASES = [
    ("景旺电子", ("景旺", "jxpcb")),
    ("赣州逸豪", ("逸豪",)),
    ("超跃科技", ("超跃", "pcb-beyond")),
    ("奥士康", ("奥士康", "askpcb")),
    ("超颖电子", ("超颖", "dynamic", "樣品需求", "样品需求")),
]

ORDER_NUMBER_PATTERNS = [
    r"(?:采购订单(?:号)?|订单号|订单编号|新订单|请购单号|PO(?: Number)?|Purchase Order)\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-_]{3,})",
    r"(GA\d{4,}(?:-\d+)?|P\d{6,}|BUM\d{6,}|OU_[A-Z0-9_]+)",
    r"\b(HJ\d{8,})\b",
]

SPEC_PATTERNS = [
    r"NY\d+[A-Z0-9]*",
    r"FR-4",
    r"覆铜板",
    r"基板",
    r"CCL",
    r"PP\b",
]

SPEC_REQUIRED = [
    re.compile(r"NY\d", re.IGNORECASE),
    re.compile(r"FR-4", re.IGNORECASE),
    re.compile(r"PP\s+(?:Halogen Free\s+)?NY", re.IGNORECASE),
    re.compile(r"CCL.*NY", re.IGNORECASE),
]


def decode_and_simplify_html(payload: bytes, charset: str = "") -> str:
    try:
        html = payload.decode(charset or "utf-8", errors="replace")
    except Exception:
        html = payload.decode("utf-8", errors="replace")
    return convert(html, "zh-cn")


def html_to_text(html: str) -> str:
    parser = _MailTextParser()
    parser.feed(html)
    parser.close()
    lines = [line.strip() for line in "".join(parser.parts).splitlines() if line.strip()]
    return "\n".join(lines)


ORDER_HINTS = ("采购订单", "订单", "po", "ga", "请购单号", "样品需求", "樣品需求", "物料", "ny")


def _forward_blocks(text: str) -> list[str]:
    parts = re.split(r"(发件人[:：]|From:)", text)
    blocks: list[str] = []
    current = ""
    for part in parts:
        if part in ("发件人：", "发件人:", "From:"):
            if current.strip():
                blocks.append(current.strip())
            current = part
        else:
            current += part
    if current.strip():
        blocks.append(current.strip())
    return blocks


def cut_latest_segment(text: str) -> str:
    for block in _forward_blocks(text):
        if "nouyatec.com" in block:
            continue
        lowered = block.lower()
        if any(hint in lowered for hint in ORDER_HINTS) or _detect_spec(block) or _detect_customer(block):
            return block.strip()
    other_markers = ["原始邮件", "-----原始邮件-----", "Sent:", "Subject:", "转发邮件"]
    positions = [index for index in (text.find(marker) for marker in other_markers) if index > 0]
    if not positions:
        return text.strip()
    return text[: min(positions)].strip()


def _detect_customer(text: str, sender: str = "") -> str:
    joined = f"{text} {sender}".lower()
    for customer, aliases in CUSTOMER_ALIASES:
        if any(alias.lower() in joined for alias in aliases):
            return customer
    return ""


def _detect_order_number(text: str) -> str:
    for pattern in ORDER_NUMBER_PATTERNS:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""


def _detect_spec(text: str) -> str:
    specs = _detect_specs(text)
    return "；".join(specs[:2])[:800]


def _detect_specs(text: str) -> list[str]:
    candidates: list[str] = []
    for line in text.splitlines():
        if any(re.search(pattern, line, flags=re.IGNORECASE) for pattern in SPEC_PATTERNS):
            cleaned = " ".join(line.split())
            if len(cleaned) < 12 or "采购组" in cleaned:
                continue
            if any(pattern.search(cleaned) for pattern in SPEC_REQUIRED) and cleaned not in candidates:
                candidates.append(cleaned)
    # Some ERP-style mails put one order line into separate visual rows. Rejoin the
    # nearby rows around the material code so it remains usable as a business hint.
    lines = [" ".join(line.split()) for line in text.splitlines() if line.strip()]
    for index, line in enumerate(lines):
        if not re.fullmatch(r"NY\d+[A-Z0-9]*", line, flags=re.IGNORECASE):
            continue
        nearby = " ".join(lines[index:index + 10])
        if "FR-4" not in nearby.upper() or nearby in candidates:
            continue
        candidates.append(nearby[:800])
    candidates.sort(key=len, reverse=True)
    return candidates[:20]


def safe_display_html(html: str, fallback_text: str = "") -> str:
    """Keep the source-mail layout while removing active or remote content."""
    source = html or fallback_text
    if not source:
        return "<p>无可展示的邮件正文。</p>"
    parser = _SafeMailHtmlParser()
    parser.feed(source[:250000])
    parser.close()
    return "".join(parser.parts) or "<p>无可展示的邮件正文。</p>"


class _MailTextParser(HTMLParser):
    """Dependency-free text extraction used by mail intake workers."""

    _SKIP_TAGS = {"script", "style", "iframe", "object", "embed", "form", "head"}
    _BREAK_TAGS = {"br", "p", "div", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.skip_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag in self._SKIP_TAGS:
            self.skip_depth += 1
        elif not self.skip_depth and tag in self._BREAK_TAGS:
            self.parts.append("\n")

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if not self.skip_depth and tag.lower() in self._BREAK_TAGS:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in self._SKIP_TAGS and self.skip_depth:
            self.skip_depth -= 1
        elif not self.skip_depth and tag in self._BREAK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self.skip_depth:
            self.parts.append(data)


class _SafeMailHtmlParser(HTMLParser):
    """Retain passive mail markup, remove executable and remote content."""

    _DROP_TAGS = {"script", "iframe", "object", "embed", "form", "base", "meta", "link", "img", "svg", "video", "audio"}
    _VOID_TAGS = {"br", "hr", "col", "input", "area", "source", "track", "wbr", "base", "meta", "link", "img", "embed"}
    _SAFE_ATTRS = {"style", "class", "id", "title", "align", "valign", "width", "height", "border", "cellpadding", "cellspacing", "bgcolor", "color", "colspan", "rowspan", "role", "dir", "lang"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.drop_depth = 0
        self.style_depth = 0

    @staticmethod
    def _safe_style(value: str) -> str:
        return re.sub(r"(?is)(url\s*\([^)]*\)|@import[^;]*;?)", "", value)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if self.drop_depth:
            self.drop_depth += 1
            return
        if tag in self._DROP_TAGS:
            if tag not in self._VOID_TAGS:
                self.drop_depth = 1
            return
        if tag == "style":
            self.style_depth += 1
        kept: list[str] = []
        for name, value in attrs:
            name = name.lower()
            if name.startswith("on") or name not in self._SAFE_ATTRS:
                continue
            clean_value = self._safe_style(value or "") if name == "style" else (value or "")
            kept.append(f' {name}="{html_lib.escape(clean_value, quote=True)}"')
        self.parts.append(f"<{tag}{''.join(kept)}>")

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() in self._DROP_TAGS:
            return
        self.handle_starttag(tag, attrs)
        if tag.lower() not in self._VOID_TAGS:
            self.handle_endtag(tag)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if self.drop_depth:
            self.drop_depth -= 1
            return
        if tag == "style" and self.style_depth:
            self.style_depth -= 1
        if tag not in self._VOID_TAGS:
            self.parts.append(f"</{tag}>")

    def handle_data(self, data: str) -> None:
        if self.drop_depth:
            return
        self.parts.append(self._safe_style(data) if self.style_depth else html_lib.escape(data))

    def handle_entityref(self, name: str) -> None:
        if not self.drop_depth:
            self.parts.append(f"&{name};")

    def handle_charref(self, name: str) -> None:
        if not self.drop_depth:
            self.parts.append(f"&#{name};")


def _detect_remark(text: str) -> str:
    remarks: list[str] = []
    for line in text.splitlines():
        if re.search(r"(?:备注|需方备注|供方备注|订单说明|要求|请)", line):
            value = re.split(r"备注\s*[:：]?\s*", line, maxsplit=1)[-1].strip()
            if value and value not in remarks:
                remarks.append(value)
        if len(remarks) >= 3:
            break
    return "；".join(remarks)[:800]


def extract_order_fields(text: str, sender: str = "") -> dict[str, str]:
    latest = cut_latest_segment(text)
    specs = _detect_specs(latest)
    return {
        "customer_name": _detect_customer(f"{latest} {text}", sender),
        "order_number": _detect_order_number(latest),
        "spec": specs[0] if specs else "",
        "specs": specs,
        "remark": _detect_remark(latest),
    }
