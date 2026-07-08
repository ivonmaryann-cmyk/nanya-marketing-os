from __future__ import annotations

import math
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .excel_utils import load_workbook_compat


FOIL_TOKEN_PATTERN = r"HS1(?:-M2)?(?:-VSP)?|HS2(?:-M2)?(?:-VSP)?|HVLP[1-4]?|RTF[1-4]?|HTE|VLP"


@dataclass
class ExtPpRule:
    excel_row: int
    sheet: str
    product: str
    glass: str
    rc_min: float | None
    rc_max: float | None
    length: int | None
    width: float | None
    price: float | None
    roll_price: float | None = None


@dataclass
class ExtCclRule:
    excel_row: int
    sheet: str
    product: str
    thickness_mm: float | None
    thickness_mil: float | None
    copper: str
    foil: str
    stack: str
    prices: dict[str, float | None]
    kind: str = ""


@dataclass
class ExtRules:
    customer_key: str
    pp_rows: list[ExtPpRule]
    ccl_rows: list[ExtCclRule]
    ccl_notes: dict[str, dict[str, Any]] = field(default_factory=dict)


@dataclass
class ExtCalcResult:
    status: str
    material_type: str
    price: float | str | None
    total: float | str | None
    width: str
    roll_length: str
    note: str
    rule_row: int | None = None
    size_column: str = ""


def load_extended_rules(customer_key: str, rule_path: str | Path) -> ExtRules:
    if customer_key == "hanyu":
        rules = _load_hanyu_rules(rule_path)
    elif customer_key == "wutong":
        rules = _load_wutong_rules(rule_path)
    elif customer_key == "eaton":
        rules = _load_eaton_rules(rule_path)
    elif customer_key == "taixing":
        rules = _load_taixing_rules(rule_path)
    elif customer_key == "aoshikang":
        rules = _load_aoshikang_rules(rule_path)
    elif customer_key == "mingyang":
        rules = _load_mingyang_rules(rule_path)
    elif customer_key == "lejian":
        rules = _load_lejian_rules(rule_path)
    elif customer_key == "guanghe":
        rules = _load_guanghe_rules(rule_path)
    elif customer_key == "shengyi":
        rules = _load_shengyi_rules(rule_path)
    elif customer_key == "zhongfu":
        rules = _load_zhongfu_rules(rule_path)
    else:
        raise ValueError(f"不支持的扩展价格计算客户：{customer_key}")
    if not rules.pp_rows and not rules.ccl_rows:
        raise ValueError("报价单未读取到有效 PP 或基板规则")
    return rules


def calculate_extended_spec(customer_key: str, spec: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    desc = _text(spec)
    if not desc:
        return ExtCalcResult("失败", "未知", "待确认", "", "", "", "客户规格为空")
    if customer_key == "taixing":
        return _calculate_taixing_spec(desc, rules, quantity=quantity)
    if customer_key == "aoshikang":
        return _calculate_aoshikang_spec(desc, rules, quantity=quantity)
    if customer_key == "mingyang":
        return _calculate_mingyang_spec(desc, rules, quantity=quantity)
    if customer_key == "lejian":
        return _calculate_lejian_spec(desc, rules, quantity=quantity)
    if customer_key == "guanghe":
        return _calculate_guanghe_spec(desc, rules, quantity=quantity)
    if customer_key == "shengyi":
        return _calculate_shengyi_spec(desc, rules, quantity=quantity)
    if customer_key == "zhongfu":
        return _calculate_zhongfu_spec(desc, rules, quantity=quantity)
    if _looks_like_pp(desc):
        return _calculate_pp(customer_key, desc, rules)
    return _calculate_ccl(customer_key, desc, rules, quantity=quantity)


def run_extended_regression(customer_key: str, rules: ExtRules, test_data_path: str | Path) -> dict:
    rows: list[dict] = []
    total = passed = failed = allowed_exception = 0
    path = Path(test_data_path)
    if not path.exists():
        return {"total": 0, "passed": 0, "failed": 0, "allowed_exception": 0, "rows": rows}
    wb = load_workbook_compat(path, data_only=True)
    for ws in wb.worksheets:
        header_row, headers = _find_header_row(ws)
        if not header_row:
            continue
        desc_col = _first_col(headers, {"客户规格", "物料长描述", "物料描述", "规格"})
        expected_col = _first_col(headers, {"新价", "新单价", "新含税价", "新价格", "单价"})
        conflict_specs = _taixing_conflict_specs(ws, header_row, desc_col, expected_col) if customer_key in {"taixing", "aoshikang"} and desc_col and expected_col else {}
        if not desc_col:
            continue
        for row_idx in range(header_row + 1, ws.max_row + 1):
            spec = _text(ws.cell(row_idx, desc_col).value)
            if not spec:
                continue
            total += 1
            result = calculate_extended_spec(customer_key, spec, rules)
            expected = ws.cell(row_idx, expected_col).value if expected_col else None
            if expected_col:
                tolerance = 0.0002 if customer_key in {"taixing", "aoshikang"} else 0.02
                ok = _result_equal(result.price, expected, tolerance=tolerance)
            else:
                ok = bool(result.note) and result.status == "成功"
            is_conflict = customer_key in {"taixing", "aoshikang"} and spec in conflict_specs and result.status == "成功"
            if ok:
                passed += 1
                status = "通过"
            elif is_conflict:
                passed += 1
                allowed_exception += 1
                status = "允许例外"
            else:
                failed += 1
                status = "失败"
            note = result.note
            if is_conflict:
                note = f"{note}；重复规格单价冲突：{conflict_specs[spec]}"
            rows.append(
                {
                    "sheet": ws.title,
                    "row": row_idx,
                    "status": status,
                    "expected_price": expected,
                    "actual_price": result.price,
                    "expected_total": "",
                    "actual_total": "",
                    "note": note,
                }
            )
    return {"total": total, "passed": passed, "failed": failed, "allowed_exception": allowed_exception, "rows": rows}


def _load_hanyu_rules(rule_path: str | Path) -> ExtRules:
    wb = load_workbook_compat(rule_path, data_only=True)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    for ws in wb.worksheets:
        for row_idx in range(1, ws.max_row + 1):
            values = [_text(ws.cell(row_idx, col).value) for col in range(1, ws.max_column + 1)]
            value_set = set(values)
            if {"Product", "mil", "Stacking", "SQ.FT"}.issubset(value_set):
                headers = _header_map(values)
                price_cols = _price_columns(values)
                for data_row in range(row_idx + 1, ws.max_row + 1):
                    product = _norm_product(ws.cell(data_row, headers.get("Product", 1)).value)
                    if not product:
                        break
                    thickness_mm = _to_float(ws.cell(data_row, headers.get("mm", 2)).value)
                    thickness_mil = _to_float(ws.cell(data_row, headers.get("mil", 3)).value)
                    copper = _norm_copper(ws.cell(data_row, headers.get("Cu", 4)).value)
                    foil = _norm_foil(ws.cell(data_row, headers.get("Cu__2", 5)).value) or "HTE"
                    stack = _norm_stack(ws.cell(data_row, headers.get("Stacking", 6)).value)
                    if thickness_mil is None or not copper or not stack:
                        continue
                    ccl_rows.append(ExtCclRule(data_row, ws.title, product, thickness_mm, thickness_mil, copper, foil, stack, _row_prices(ws, data_row, price_cols)))
            if {"Products", "Glass type", "Resin Content", "Length (m)"}.issubset(value_set):
                headers = _header_map(values)
                per_m_col = _first_header_col(values, {"Per M"})
                per_roll_col = _first_header_col(values, {"Per Roll"})
                for data_row in range(row_idx + 1, ws.max_row + 1):
                    product = _norm_product(ws.cell(data_row, headers.get("Products", 1)).value)
                    glasses = _norm_glasses(ws.cell(data_row, headers.get("Glass type", 2)).value)
                    rc_min, rc_max = _parse_rc_range(ws.cell(data_row, headers.get("Resin Content", 3)).value)
                    length = _length_int(ws.cell(data_row, headers.get("Length (m)", 4)).value)
                    price = _to_float(ws.cell(data_row, per_m_col).value) if per_m_col else None
                    roll_price = _to_float(ws.cell(data_row, per_roll_col).value) if per_roll_col else None
                    if not product:
                        break
                    if glasses and rc_min is not None and price is not None:
                        for glass in glasses:
                            pp_rows.append(ExtPpRule(data_row, ws.title, product, glass, rc_min, rc_max, length, None, price, roll_price))
    return ExtRules("hanyu", pp_rows, ccl_rows)


def _load_wutong_rules(rule_path: str | Path) -> ExtRules:
    wb = load_workbook_compat(rule_path, data_only=True)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    for ws in wb.worksheets:
        for row_idx in range(1, ws.max_row + 1):
            values = [_text(ws.cell(row_idx, col).value) for col in range(1, ws.max_column + 1)]
            joined = "|".join(values)
            if "产品型号" in values and "厚度" in values and "叠构" in values and "SQ.FT" in values:
                price_cols = _price_columns(values)
                for data_row in range(row_idx + 1, ws.max_row + 1):
                    product = _norm_product(ws.cell(data_row, 1).value)
                    if not product:
                        break
                    thickness_mm = _to_float(ws.cell(data_row, 2).value)
                    stack = _norm_stack(ws.cell(data_row, 3).value)
                    copper = _norm_copper(ws.cell(data_row, 4).value)
                    if thickness_mm is None or not stack or not copper:
                        continue
                    ccl_rows.append(ExtCclRule(data_row, ws.title, product, thickness_mm, None, copper, "", stack, _row_prices(ws, data_row, price_cols)))
            if "产品型号" in values and "布种" in values and ("RC范围" in values or "RC" in joined):
                for data_row in range(row_idx + 1, ws.max_row + 1):
                    product = _norm_product(ws.cell(data_row, 1).value)
                    if not product:
                        break
                    glasses = _norm_glasses(ws.cell(data_row, 2).value)
                    length = _length_int(ws.cell(data_row, 6).value)
                    rc_min, rc_max = _parse_rc_range(ws.cell(data_row, 8).value)
                    price = _to_float(ws.cell(data_row, 10).value)
                    if glasses and rc_min is not None and price is not None:
                        for glass in glasses:
                            pp_rows.append(ExtPpRule(data_row, ws.title, product, glass, rc_min, rc_max, length, None, price))
    return ExtRules("wutong", pp_rows, ccl_rows)


def _load_eaton_rules(rule_path: str | Path) -> ExtRules:
    wb = load_workbook_compat(rule_path, data_only=True)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    for ws in wb.worksheets:
        for row_idx in range(1, ws.max_row + 1):
            values = [_text(ws.cell(row_idx, col).value) for col in range(1, ws.max_column + 1)]
            value_set = set(values)
            if {"Products", "Glass type", "Resin Content", "Length (m)"}.issubset(value_set):
                headers = _header_map(values)
                per_m_col = _first_header_col(values, {"每米单价", "Per M", "M\n新价格"})
                per_roll_col = _first_header_col(values, {"每卷单价", "Per Roll"})
                for data_row in range(row_idx + 1, ws.max_row + 1):
                    product = _norm_product(ws.cell(data_row, headers.get("Products", 1)).value)
                    if not product:
                        break
                    glasses = _norm_glasses(ws.cell(data_row, headers.get("Glass type", 2)).value)
                    rc_min, rc_max = _parse_rc_range(ws.cell(data_row, headers.get("Resin Content", 3)).value)
                    length = _length_int(ws.cell(data_row, headers.get("Length (m)", 4)).value)
                    width = _to_float(ws.cell(data_row, headers.get("Width (inch)", 5)).value)
                    price = _to_float(ws.cell(data_row, per_m_col).value) if per_m_col else None
                    roll_price = _to_float(ws.cell(data_row, per_roll_col).value) if per_roll_col else None
                    if glasses and rc_min is not None and (price is not None or roll_price is not None):
                        for glass in glasses:
                            pp_rows.append(ExtPpRule(data_row, ws.title, product, glass, rc_min, rc_max, length, width, price, roll_price))
            if _looks_like_eaton_ccl_header(values):
                price_cols = _price_columns(values)
                product_col = _find_header_contains(values, {"Products", "产品类别", "型号"}) or 1
                mm_col = _find_header_contains(values, {"厚度mm", "厚(mm", "厚(mm）"})
                mil_col = _find_header_contains(values, {"mil"})
                stack_col = _find_header_contains(values, {"组合结构", "结构"}) or 4
                for data_row in range(row_idx + 1, ws.max_row + 1):
                    product = _norm_product(ws.cell(data_row, product_col).value)
                    if not product:
                        break
                    thickness_mm = _to_float(ws.cell(data_row, mm_col).value) if mm_col else None
                    thickness_mil = _to_float(ws.cell(data_row, mil_col).value) if mil_col else None
                    stack = _norm_taixing_stack(ws.cell(data_row, stack_col).value)
                    copper, foil = _eaton_copper_and_foil(ws, data_row)
                    foil = foil or "HTE"
                    if thickness_mil is None or not stack or not copper:
                        continue
                    ccl_rows.append(ExtCclRule(data_row, ws.title, product, thickness_mm, thickness_mil, copper, foil, stack, _row_prices(ws, data_row, price_cols)))
    return ExtRules("eaton", pp_rows, ccl_rows)


def _load_taixing_rules(rule_path: str | Path) -> ExtRules:
    wb = load_workbook_compat(rule_path, data_only=True)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    ccl_notes: dict[str, dict[str, Any]] = {}
    for ws in wb.worksheets:
        for row_idx in range(1, ws.max_row + 1):
            values = [_text(ws.cell(row_idx, col).value) for col in range(1, ws.max_column + 1)]
            value_set = set(values)
            if {"Products", "Glass type", "Resin Content", "Length (m)"}.issubset(value_set):
                headers = _header_map(values)
                per_m_col = _first_header_col(values, {"每米单价", "Per Meter", "Per M"})
                per_roll_col = _first_header_col(values, {"每卷单价", "Per Roll"})
                width_col = _first_header_col(values, {"Width (inch)"})
                for data_row in range(row_idx + 1, ws.max_row + 1):
                    product = _norm_product(ws.cell(data_row, headers.get("Products", 1)).value)
                    if not product:
                        break
                    glasses = _norm_taixing_glasses(ws.cell(data_row, headers.get("Glass type", 2)).value)
                    rc_min, rc_max = _parse_rc_range(ws.cell(data_row, headers.get("Resin Content", 3)).value)
                    length = _length_int(ws.cell(data_row, headers.get("Length (m)", 4)).value)
                    width = _to_float(ws.cell(data_row, width_col).value) if width_col else None
                    price = _to_float(ws.cell(data_row, per_m_col).value) if per_m_col else None
                    roll_price = _to_float(ws.cell(data_row, per_roll_col).value) if per_roll_col else None
                    if price is None and roll_price is not None and length:
                        price = roll_price / length
                    if glasses and rc_min is not None and price is not None:
                        for glass in glasses:
                            pp_rows.append(ExtPpRule(data_row, ws.title, product, glass, rc_min, rc_max, length, width, price, roll_price))
            if _looks_like_taixing_ccl_header(values):
                price_cols = _taixing_ccl_price_columns(ws, row_idx, values)
                ccl_notes[ws.title] = _parse_taixing_ccl_sheet_notes(ws)
                headers = _header_map(values)
                product_col = _find_header_contains(values, {"Product", "产品类别"}) or 1
                mm_col = _find_header_contains(values, {"mm"})
                mil_col = _find_header_contains(values, {"mil"})
                stack_col = _find_header_contains(values, {"结构", "组合"}) or 4
                for data_row in range(row_idx + 1, ws.max_row + 1):
                    product = _norm_product(ws.cell(data_row, product_col).value)
                    if not product:
                        break
                    thickness_mm = _to_float(ws.cell(data_row, mm_col).value) if mm_col else None
                    thickness_mil = _to_float(ws.cell(data_row, mil_col).value) if mil_col else None
                    stack = _norm_stack(ws.cell(data_row, stack_col).value)
                    if thickness_mm is None or not stack:
                        continue
                    ccl_rows.append(
                        ExtCclRule(
                            data_row,
                            ws.title,
                            product,
                            thickness_mm,
                            thickness_mil,
                            "",
                            "",
                            stack,
                            _row_prices(ws, data_row, price_cols),
                        )
                    )
    return ExtRules("taixing", pp_rows, ccl_rows, ccl_notes)


def _load_aoshikang_rules(rule_path: str | Path) -> ExtRules:
    wb = load_workbook_compat(rule_path, data_only=True)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    product_meta: dict[str, dict[str, Any]] = {}
    for ws in wb.worksheets:
        if ws.title.strip().upper() == "PP":
            _load_aoshikang_pp_sheet(ws, pp_rows)
            continue
        header_rows: list[tuple[int, dict[str, Any]]] = []
        for row_idx in range(1, ws.max_row + 1):
            values = [_text(ws.cell(row_idx, col).value) for col in range(1, ws.max_column + 1)]
            info = _aoshikang_ccl_header_info(values)
            if info:
                header_rows.append((row_idx, info))
        for index, (header_row, info) in enumerate(header_rows):
            next_header = header_rows[index + 1][0] if index + 1 < len(header_rows) else ws.max_row + 1
            products = _aoshikang_product_aliases(info["product"])
            price_cols = info["price_cols"]
            if not products or not price_cols:
                continue
            for product in products:
                product_meta.setdefault(product, _aoshikang_product_meta(product, bool(info.get("foil_col"))))
            for data_row in range(header_row + 1, next_header):
                thickness = _to_float(ws.cell(data_row, 1).value)
                copper = _aoshikang_norm_copper(ws.cell(data_row, 2).value)
                raw_stack = ws.cell(data_row, 3).value
                stack = _aoshikang_norm_stack(raw_stack)
                if thickness is None or not copper or not stack:
                    continue
                kind_text = _text(ws.cell(data_row, info["kind_col"]).value) if info.get("kind_col") else ""
                foil = ""
                if info.get("foil_col"):
                    foil = _aoshikang_norm_foil(ws.cell(data_row, info["foil_col"]).value)
                if not foil:
                    foil = _aoshikang_norm_foil(f"{raw_stack} {kind_text}")
                prices = _row_prices(ws, data_row, price_cols)
                tax_price = _to_float(ws.cell(data_row, info["tax_col"]).value) if info.get("tax_col") else None
                if tax_price is not None:
                    prices["TAX"] = tax_price
                if not prices:
                    continue
                for product in products:
                    ccl_rows.append(ExtCclRule(data_row, ws.title, product, thickness, None, copper, foil, stack, prices, kind_text))
    return ExtRules("aoshikang", pp_rows, ccl_rows, {"products": product_meta})


def _load_mingyang_rules(rule_path: str | Path) -> ExtRules:
    wb = load_workbook_compat(rule_path, data_only=True)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    for ws in wb.worksheets:
        title = ws.title.strip()
        for row_idx in range(1, ws.max_row + 1):
            values = [_text(ws.cell(row_idx, col).value) for col in range(1, min(ws.max_column, 20) + 1)]
            if title in {"通用PP", "高速PP"} and _mingyang_is_pp_header(values):
                _load_mingyang_pp_rows(ws, row_idx, values, pp_rows)
                break
            if title in {"通用CCL", "通用CCL 高速"} and _mingyang_is_ccl_header(values):
                _load_mingyang_ccl_rows(ws, row_idx, values, ccl_rows)
                break
    return ExtRules("mingyang", pp_rows, ccl_rows)


def _mingyang_is_pp_header(values: list[str]) -> bool:
    value_set = set(values)
    return bool({"布种", "RC含量", "长度", "宽幅"}.issubset(value_set) and ({"产品型号", "新价格"}.issubset(value_set) or {"胶系", "单价"}.issubset(value_set)))


def _load_mingyang_pp_rows(ws, header_row: int, values: list[str], pp_rows: list[ExtPpRule]) -> None:
    headers = _header_map(values)
    product_col = headers.get("产品型号") or headers.get("胶系") or 2
    glass_col = headers.get("布种") or 3
    rc_col = headers.get("RC含量") or 4
    length_col = headers.get("长度") or 5
    width_col = headers.get("宽幅") or 6
    price_col = headers.get("新价格") or headers.get("单价") or 7
    for data_row in range(header_row + 1, ws.max_row + 1):
        first = _text(ws.cell(data_row, 1).value)
        if first.startswith("说明"):
            break
        product = _norm_product(ws.cell(data_row, product_col).value)
        if not product:
            continue
        glasses = _norm_glasses(ws.cell(data_row, glass_col).value)
        rc_min, rc_max = _parse_rc_range(ws.cell(data_row, rc_col).value)
        length = _length_int(ws.cell(data_row, length_col).value)
        width = _to_float(ws.cell(data_row, width_col).value)
        price = _to_float(ws.cell(data_row, price_col).value)
        if glasses and rc_min is not None and price is not None:
            for glass in glasses:
                pp_rows.append(ExtPpRule(data_row, ws.title, product, glass, rc_min, rc_max, length, width, price))


def _mingyang_is_ccl_header(values: list[str]) -> bool:
    value_set = set(values)
    return {"产品型号", "板厚", "铜箔", "铜箔特性", "配料结构"}.issubset(value_set) and any(_price_key_from_label(value) in {"37", "41", "43"} for value in values)


def _load_mingyang_ccl_rows(ws, header_row: int, values: list[str], ccl_rows: list[ExtCclRule]) -> None:
    headers = _header_map(values)
    product_col = headers.get("产品型号") or 2
    thickness_col = headers.get("板厚") or 4
    copper_col = headers.get("铜箔") or 6
    foil_col = headers.get("铜箔特性") or 7
    stack_col = headers.get("配料结构") or 8
    price_cols = _mingyang_price_columns(values)
    for data_row in range(header_row + 1, ws.max_row + 1):
        first = _text(ws.cell(data_row, 1).value)
        if first.startswith("说明"):
            break
        product = _norm_product(ws.cell(data_row, product_col).value)
        thickness_values = _parse_mingyang_thickness_values(ws.cell(data_row, thickness_col).value)
        copper = _norm_copper(ws.cell(data_row, copper_col).value)
        foil = _norm_foil(ws.cell(data_row, foil_col).value)
        stack = _norm_stack(ws.cell(data_row, stack_col).value)
        if product and thickness_values and copper and stack:
            prices = _row_prices(ws, data_row, price_cols)
            if prices:
                for thickness_mm in thickness_values:
                    ccl_rows.append(ExtCclRule(data_row, ws.title, product, thickness_mm, None, copper, foil, stack, prices))


def _parse_mingyang_thickness_values(value: Any) -> list[float]:
    text = _text(value).replace(",", "")
    if not text:
        return []
    values: list[float] = []
    seen: set[float] = set()
    for match in re.findall(r"-?\d+(?:\.\d+)?", text):
        thickness = float(match)
        if thickness in seen:
            continue
        seen.add(thickness)
        values.append(thickness)
    return values


def _mingyang_price_columns(headers: list[str]) -> dict[int, str]:
    price_cols: dict[int, str] = {}
    for idx, header in enumerate(headers, start=1):
        normalized = _text(header).replace("\n", "").replace(" ", "")
        if not normalized or "旧" in normalized or "调整" in normalized or "备注" in normalized:
            continue
        key = _price_key_from_label(normalized)
        if key:
            price_cols[idx] = key
    return price_cols

def _load_lejian_rules(rule_path: str | Path) -> ExtRules:
    sheets = _workbook_value_sheets(rule_path)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    for sheet_name, rows in sheets:
        for row_index, values in enumerate(rows, start=1):
            normalized = [_text(value).replace("\n", "").replace(" ", "") for value in values]
            if {"型号", "铜箔OZ", "布种叠构"}.issubset(set(normalized)):
                _load_lejian_ccl_rows(sheet_name, rows, row_index, normalized, ccl_rows)
            if "布种" in normalized and "含量%" in normalized:
                _load_lejian_pp_rows(sheet_name, rows, row_index, normalized, pp_rows)
    return ExtRules("lejian", pp_rows, ccl_rows)


def _load_lejian_ccl_rows(
    sheet_name: str,
    rows: list[list[Any]],
    header_row: int,
    headers: list[str],
    ccl_rows: list[ExtCclRule],
) -> None:
    header_map = _header_map(headers)
    product_col = header_map.get("型号") or 3
    thickness_col = header_map.get("含铜厚度mm") or header_map.get("含铜厚度") or 4
    copper_col = header_map.get("铜箔OZ") or 6
    foil_top_col = header_map.get("上铜箔类型") or 7
    foil_bottom_col = header_map.get("下铜箔类型") or 8
    stack_col = header_map.get("布种叠构") or 9
    price_cols = _lejian_ccl_price_columns(headers)
    for data_row in range(header_row + 1, len(rows) + 1):
        values = rows[data_row - 1]
        product = _norm_product(_row_value(values, product_col))
        if not product:
            continue
        if _text(_row_value(values, 3)) == "布种":
            break
        copper = _norm_copper(_row_value(values, copper_col))
        foil = _norm_foil(_row_value(values, foil_top_col)) or _norm_foil(_row_value(values, foil_bottom_col)) or "HTE"
        stack = _norm_stack(_row_value(values, stack_col))
        prices = {key: _to_float(_row_value(values, col)) for col, key in price_cols.items()}
        prices = {key: value for key, value in prices.items() if value is not None}
        thickness_values = _parse_lejian_thickness_values(_row_value(values, thickness_col))
        if not copper or not prices or not thickness_values:
            continue
        for thickness_mm in thickness_values:
            ccl_rows.append(ExtCclRule(data_row, sheet_name, product, thickness_mm, None, copper, foil, stack, prices))


def _load_lejian_pp_rows(
    sheet_name: str,
    rows: list[list[Any]],
    header_row: int,
    headers: list[str],
    pp_rows: list[ExtPpRule],
) -> None:
    header_map = _header_map(headers)
    glass_col = header_map.get("布种") or 3
    rc_col = header_map.get("含量%") or 4
    product_cols: list[tuple[int, str]] = []
    for col_idx, header in enumerate(headers, start=1):
        product = _norm_product(header)
        if product.startswith("NY"):
            product_cols.append((col_idx, product))
    for data_row in range(header_row + 1, len(rows) + 1):
        values = rows[data_row - 1]
        glass = _norm_glass(_row_value(values, glass_col))
        rc_min, rc_max = _parse_rc_range(_row_value(values, rc_col))
        if not glass or rc_min is None or rc_max is None:
            continue
        for col_idx, product in product_cols:
            price = _to_float(_row_value(values, col_idx))
            if price is not None:
                pp_rows.append(ExtPpRule(data_row, sheet_name, product, glass, rc_min, rc_max, None, 49.5, price))


def _calculate_lejian_spec(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    if _looks_like_pp(desc) or "半固化片" in desc:
        return _calculate_lejian_pp(desc, rules)
    return _calculate_lejian_ccl(desc, rules, quantity=quantity)


def _calculate_lejian_pp(desc: str, rules: ExtRules) -> ExtCalcResult:
    product = _extract_lejian_product(desc)
    glass = _extract_glass(desc)
    rc = _extract_lejian_rc(desc)
    if not product or not glass or rc is None:
        return ExtCalcResult("失败", "PP", "待确认", "", "", "", "乐健PP规格缺少型号、布种或RC")
    product_norm = _norm_product(product)
    base_product = product_norm[:-1] if product_norm.endswith("P") else product_norm
    products = {product_norm, base_product, f"{base_product}P"}
    matches = [
        row
        for row in rules.pp_rows
        if row.product in products
        and row.glass == glass
        and row.rc_min is not None
        and row.rc_max is not None
        and row.rc_min - 0.001 <= rc <= row.rc_max + 0.001
    ]
    if not matches:
        return ExtCalcResult("失败", "PP", "待确认", "", "", "", "未命中乐健PP报价：型号、布种或RC不匹配")
    best = sorted(matches, key=lambda row: (0 if row.product == base_product else 1, row.excel_row))[0]
    if best.price is None:
        return ExtCalcResult("失败", "PP", "待确认", "", "", "", "命中乐健PP报价行但单价为空")
    price = _round_money(best.price)
    note = f"命中乐健PP报价 Sheet {best.sheet} 第{best.excel_row}行，布种={glass}，RC={rc:g}，按旧报价表原价={price:.2f}"
    return ExtCalcResult("成功", "PP", price, "", _fmt_width(best.width), "", note, best.excel_row, best.sheet)


def _calculate_lejian_ccl(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    product = _extract_lejian_product(desc)
    thickness_mm = _extract_lejian_thickness_mm(desc)
    copper = _extract_copper(desc)
    length_in, width_in = _extract_lejian_size(desc)
    if not product or thickness_mm is None or not copper or length_in is None or width_in is None:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "乐健CCL规格缺少型号、厚度、铜厚或尺寸")
    product_norm = _norm_product(product)
    candidates = [
        row
        for row in rules.ccl_rows
        if row.product == product_norm
        and _lejian_thickness_matches(row, thickness_mm)
        and (row.copper == copper or row.copper == _reverse_copper(copper))
    ]
    if not candidates:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "未命中乐健CCL报价：型号、厚度或铜厚不匹配")
    stack = _extract_stack(desc)
    if stack:
        stack_matches = [row for row in candidates if row.stack == stack]
        if stack_matches:
            candidates = stack_matches
    row = sorted(candidates, key=lambda item: item.excel_row)[0]
    price_result = _lejian_ccl_size_price(row, length_in, width_in)
    if not price_result["ok"]:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", price_result["reason"])
    price = _round_money(price_result["price"])
    total = _calc_total(quantity, price)
    note = (
        f"命中乐健CCL报价 Sheet {row.sheet} 第{row.excel_row}行，厚度={row.thickness_mm:g}，"
        f"铜厚={row.copper}，尺寸列={price_result['label']}，公式={price_result['formula']}"
    )
    return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, price_result["label"])


def _lejian_ccl_size_price(row: ExtCclRule, length_in: float, width_in: float) -> dict[str, Any]:
    dim = _lejian_sheet_width(length_in, width_in)
    for target, key in [(37, "37"), (41, "41"), (43, "43")]:
        if abs(dim - target) <= 0.8 and row.prices.get(key) is not None:
            price = float(row.prices[key])
            return {"ok": True, "price": price, "label": f"{key}*49", "formula": f"{price:.6g}"}
    for target, source_key, factor in [(82, "41", 2), (86, "43", 2)]:
        if abs(dim - target) <= 0.8 and row.prices.get(source_key) is not None:
            base = float(row.prices[source_key])
            return {"ok": True, "price": base * factor, "label": f"{target}*49", "formula": f"{base:.6g}*{factor:g}"}
    return {"ok": False, "reason": f"乐健CCL未找到可用尺寸报价：{length_in:g}*{width_in:g}"}


def _extract_lejian_product(desc: str) -> str:
    matches = re.findall(r"NY\s*-?\s*[A-Z]?\d{3,4}[A-Z0-9]*P?", desc, re.I)
    return _norm_product(matches[-1]) if matches else ""


def _extract_lejian_rc(desc: str) -> float | None:
    match = re.search(r"RC\s*[:：]?\s*(\d+(?:\.\d+)?)", desc, re.I)
    if not match:
        return None
    value = float(match.group(1))
    return value * 100 if value <= 1 else value


def _extract_lejian_thickness_mm(desc: str) -> float | None:
    match = re.search(r"FR-?\s*4\s+(\d+(?:\.\d+)?)", desc, re.I)
    if match:
        return float(match.group(1))
    match = re.search(r"\b(\d+(?:\.\d+)?)\s*(?:\(|（)?\s*(?:含铜|不含铜)", desc, re.I)
    if match:
        return float(match.group(1))
    return _extract_thickness_mm(desc)


def _extract_lejian_size(desc: str) -> tuple[float | None, float | None]:
    candidates: list[tuple[float, float]] = []
    for match in re.finditer(r"(\d+(?:\.\d+)?)\s*[*xX×]\s*(\d+(?:\.\d+)?)", desc):
        left = float(match.group(1))
        right = float(match.group(2))
        if left <= 120 and right <= 120 and max(left, right) >= 30:
            candidates.append((left, right))
    if candidates:
        return candidates[-1]
    return _extract_size(desc, ignore_decimal=False)


def _lejian_sheet_width(length_in: float, width_in: float) -> float:
    if abs(length_in - 49) <= 1.0:
        return width_in
    if abs(width_in - 49) <= 1.0:
        return length_in
    return max(length_in, width_in)


def _lejian_ccl_price_columns(headers: list[str]) -> dict[int, str]:
    price_cols: dict[int, str] = {}
    for idx, header in enumerate(headers, start=1):
        normalized = _text(header).replace(" ", "")
        if "SF" in normalized:
            price_cols[idx] = "SF"
        elif "37*49" in normalized:
            price_cols[idx] = "37"
        elif "41*49" in normalized:
            price_cols[idx] = "41"
        elif "43*49" in normalized:
            price_cols[idx] = "43"
    return price_cols


def _parse_lejian_thickness_values(value: Any) -> list[float]:
    text = _text(value).replace(",", "")
    values: list[float] = []
    seen: set[float] = set()
    for item in re.findall(r"\d+(?:\.\d+)?", text):
        number = float(item)
        if 0 < number < 10 and number not in seen:
            seen.add(number)
            values.append(number)
    return values


def _lejian_thickness_matches(row: ExtCclRule, thickness_mm: float) -> bool:
    return row.thickness_mm is not None and abs(row.thickness_mm - thickness_mm) <= 0.012


def _load_shengyi_rules(rule_path: str | Path) -> ExtRules:
    wb = load_workbook_compat(rule_path, data_only=True)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    for ws in wb.worksheets:
        max_col = min(ws.max_column, 24)
        for row_idx in range(1, ws.max_row + 1):
            values = [_text(ws.cell(row_idx, col).value).replace("\n", "").replace("\u00a0", " ").strip() for col in range(1, max_col + 1)]
            if _shengyi_is_ccl_header(values):
                _load_shengyi_ccl_rows(ws, row_idx, values, ccl_rows)
            elif _shengyi_is_pp_header(values):
                _load_shengyi_pp_rows(ws, row_idx, values, pp_rows)
    return ExtRules("shengyi", pp_rows, ccl_rows)


def _load_shengyi_ccl_rows(ws, header_row: int, headers: list[str], ccl_rows: list[ExtCclRule]) -> None:
    product_col = _find_header_contains(headers, {"Type"}) or 1
    mm_col = _find_header_contains(headers, {"Thickness (mm)", "Thickness(mm)"}) or 2
    mil_col = _find_header_contains(headers, {"Thickness (mil)", "Thickness(mil)"}) or 3
    copper_col = _find_header_contains(headers, {"Copper"}) or 4
    foil_col = _find_header_contains(headers, {"copper type"}) or 5
    stack_col = _find_header_contains(headers, {"Structure"}) or 6
    price_cols = _shengyi_ccl_price_columns(headers)
    product = _shengyi_product_from_text(ws.title)
    started = False
    blank_streak = 0
    for data_row in range(header_row + 1, ws.max_row + 1):
        row_values = [_text(ws.cell(data_row, col).value) for col in range(1, min(ws.max_column, 24) + 1)]
        if _shengyi_is_pp_header(row_values) or any("说明" in value for value in row_values):
            break
        if not any(row_values):
            if started:
                blank_streak += 1
                if blank_streak >= 2:
                    break
            continue
        blank_streak = 0
        row_product = _shengyi_product_from_text(_row_value(row_values, product_col))
        if row_product:
            product = row_product
        thickness_mm = _to_float(_row_value(row_values, mm_col))
        thickness_mil = _to_float(_row_value(row_values, mil_col))
        copper = _norm_copper(_row_value(row_values, copper_col))
        foil = _shengyi_norm_foil(_row_value(row_values, foil_col))
        stack = _norm_stack(_row_value(row_values, stack_col))
        prices = _row_prices(ws, data_row, price_cols)
        prices = {key: value for key, value in prices.items() if value is not None}
        if not product or thickness_mm is None or thickness_mil is None or not copper or not stack or not prices:
            continue
        started = True
        ccl_rows.append(ExtCclRule(data_row, ws.title, product, thickness_mm, thickness_mil, copper, foil, stack, prices))


def _load_shengyi_pp_rows(ws, header_row: int, headers: list[str], pp_rows: list[ExtPpRule]) -> None:
    product_col = _find_header_contains(headers, {"Type"}) or 1
    glass_col = _find_header_contains(headers, {"Glass"}) or 2
    rc_col = _find_header_contains(headers, {"R/C", "RC"}) or 3
    length_col = _find_header_contains(headers, {"Length"}) or 4
    width_col = _find_header_contains(headers, {"Width"}) or 5
    price_col = _shengyi_pp_price_col(headers)
    sf_col = _find_header_contains(headers, {"SF/RMB"})
    product = _shengyi_product_from_text(ws.title)
    started = False
    blank_streak = 0
    for data_row in range(header_row + 1, ws.max_row + 1):
        row_values = [_text(ws.cell(data_row, col).value) for col in range(1, min(ws.max_column, 16) + 1)]
        if any("说明" in value for value in row_values):
            break
        if not any(row_values):
            if started:
                blank_streak += 1
                if blank_streak >= 2:
                    break
            continue
        blank_streak = 0
        row_product = _shengyi_product_from_text(_row_value(row_values, product_col))
        if row_product:
            product = row_product
        glasses = _norm_glasses(_row_value(row_values, glass_col))
        rc_min, rc_max = _parse_rc_range(_row_value(row_values, rc_col))
        length = _length_int(_row_value(row_values, length_col))
        width = _to_float(_row_value(row_values, width_col))
        price = _to_float(ws.cell(data_row, price_col).value) if price_col else None
        sf_price = _to_float(ws.cell(data_row, sf_col).value) if sf_col else None
        if price is None and sf_price is not None:
            price = _round_money(sf_price * 13.124) if sf_price is not None else None
        if not product or not glasses or rc_min is None or rc_max is None or price is None:
            continue
        started = True
        for glass in glasses:
            pp_rows.append(ExtPpRule(data_row, ws.title, product, glass, rc_min, rc_max, length, width, price, sf_price))


def _calculate_shengyi_spec(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    if _shengyi_looks_like_pp(desc):
        return _calculate_shengyi_pp(desc, rules)
    return _calculate_shengyi_ccl(desc, rules, quantity=quantity)


def _calculate_shengyi_pp(desc: str, rules: ExtRules) -> ExtCalcResult:
    product = _shengyi_product_from_text(desc)
    glass = _extract_glass(desc)
    rc = _shengyi_extract_rc(desc)
    length = _extract_length(desc, floor_value=True)
    pp_piece = _shengyi_extract_pp_piece(desc)
    width = _shengyi_extract_pp_width(desc) or _extract_width(desc) or (pp_piece["mother_width"] if pp_piece else None)
    if not product or not glass or rc is None:
        return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width), _shengyi_fmt_roll_length(length), "生益PP规格缺少型号、玻布或RC")
    products = _shengyi_product_aliases(product, for_pp=True)
    matches = [
        row
        for row in rules.pp_rows
        if row.product in products
        and row.glass == glass
        and row.rc_min is not None
        and row.rc_max is not None
        and row.rc_min - 0.001 <= rc <= row.rc_max + 0.001
        and (length is None or row.length is None or row.length == length)
        and (width is None or row.width is None or abs(row.width - width) <= 0.8)
    ]
    if not matches:
        return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width), _shengyi_fmt_roll_length(length), "未命中生益PP报价：型号、玻布、RC、卷长或宽度不匹配")
    best = sorted(matches, key=lambda row: (0 if row.product == product else 1, row.excel_row))[0]
    if pp_piece:
        piece_price = _shengyi_piece_formula_price(
            pp_piece["radial"],
            pp_piece["latitudinal"],
            pp_piece["mother_width"] or best.width,
            best.roll_price,
        )
        if not piece_price["ok"]:
            return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width or best.width), _shengyi_fmt_roll_length(length or best.length), piece_price["reason"])
        price = _round_money(piece_price["price"])
        note = (
            f"命中生益PP报价 Sheet {best.sheet} 第{best.excel_row}行，"
            f"型号={best.product}，玻布={glass}，RC={rc:g}，小片公式={piece_price['formula']}"
        )
        return ExtCalcResult("成功", "PP", price, "", _fmt_width(width or best.width), _shengyi_fmt_roll_length(length or best.length), note, best.excel_row, best.sheet)
    price = _round_money(best.price)
    note = (
        f"命中生益PP报价 Sheet {best.sheet} 第{best.excel_row}行，"
        f"型号={best.product}，玻布={glass}，RC={rc:g}，每米价={price:.2f}"
    )
    return ExtCalcResult("成功", "PP", price, "", _fmt_width(width or best.width), _shengyi_fmt_roll_length(length or best.length), note, best.excel_row, best.sheet)


def _calculate_shengyi_ccl(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    product = _shengyi_product_from_text(desc)
    thickness_mm = _shengyi_extract_thickness_mm(desc)
    thickness_mil = thickness_mm / 0.0254 if thickness_mm is not None else _extract_thickness_mil(desc, product)
    copper = _extract_copper(desc)
    foil = _shengyi_norm_foil(desc)
    stack = _extract_stack(desc)
    length_in, width_in = _shengyi_extract_size(desc)
    if not product or thickness_mm is None or not copper or not stack or length_in is None or width_in is None:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "生益CCL规格缺少型号、厚度、铜厚、叠构或尺寸")
    products = _shengyi_product_aliases(product, for_pp=False)
    product_rows = [
        row
        for row in rules.ccl_rows
        if row.product in products
        and _shengyi_thickness_matches(row, thickness_mil, thickness_mm)
        and row.stack == stack
        and row.copper in _shengyi_copper_aliases(copper)
    ]
    if not product_rows:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "未命中生益CCL报价：型号、厚度、铜厚或叠构不匹配")
    last_reason = ""
    for row in sorted(product_rows, key=lambda item: (0 if item.product == product else 1, _shengyi_thickness_distance(item, thickness_mil, thickness_mm), 0 if item.foil == foil else 1, item.excel_row)):
        adjusted = _shengyi_adjusted_sf(row, foil, copper)
        if not adjusted["ok"]:
            last_reason = adjusted["reason"]
            continue
        price_result = _shengyi_ccl_size_price(row, length_in, width_in, adjusted["sf"])
        if price_result["ok"]:
            price = _round_money(price_result["price"])
            total = _calc_total(quantity, price)
            note = (
                f"命中生益CCL报价 Sheet {row.sheet} 第{row.excel_row}行，"
                f"型号={row.product}，厚度={row.thickness_mm:g}mm，铜厚={row.copper}，铜箔={foil or row.foil or '默认'}，"
                f"叠构={row.stack}，尺寸列={price_result['label']}，公式={price_result['formula']}{adjusted['note']}"
            )
            return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, price_result["label"])
        last_reason = price_result["reason"]
    return ExtCalcResult("失败", "CCL", "待确认", "", "", "", last_reason or "生益CCL报价行找到，但尺寸或铜箔规则未匹配")


def _shengyi_is_ccl_header(values: list[str]) -> bool:
    compact = {_text(value).replace(" ", "").replace("\u00a0", "").upper() for value in values}
    return "TYPE" in compact and "THICKNESS(MIL)" in compact and "PERSF" in compact and any("36" in value and "48" in value for value in compact)


def _shengyi_looks_like_pp(desc: str) -> bool:
    upper = desc.upper()
    return "半固化片" in desc or ("RC" in upper and not ("覆铜" in desc or "基板" in desc))


def _shengyi_fmt_roll_length(value: int | None) -> str:
    return f"{value}m" if value is not None else ""


def _shengyi_is_pp_header(values: list[str]) -> bool:
    compact = {_text(value).replace(" ", "").replace("\u00a0", "").upper() for value in values}
    return "TYPE" in compact and "GLASS" in compact and ("R/C" in compact or "RC" in compact) and any("RMB/M" in value for value in compact)


def _shengyi_ccl_price_columns(headers: list[str]) -> dict[int, str]:
    price_cols: dict[int, str] = {}
    for idx, header in enumerate(headers, start=1):
        normalized = _text(header).upper().replace('"', "").replace(" ", "").replace("\u00a0", "")
        if not normalized:
            continue
        if normalized == "PERSF":
            price_cols[idx] = "SF"
        elif "36" in normalized and "48" in normalized:
            price_cols[idx] = "36*48"
        elif "40" in normalized and "48" in normalized:
            price_cols[idx] = "40*48"
        elif "42" in normalized and "48" in normalized:
            price_cols[idx] = "42*48"
        elif "37*43" in normalized:
            price_cols[idx] = "37*43"
        else:
            key = _price_key_from_label(header)
            if key:
                price_cols[idx] = key
    return price_cols


def _shengyi_pp_price_col(headers: list[str]) -> int | None:
    for idx, header in enumerate(headers, start=1):
        normalized = _text(header).upper().replace(" ", "").replace("\u00a0", "")
        if "RMB/M" in normalized:
            return idx
    return None


def _shengyi_product_from_text(value: Any) -> str:
    text = _text(value)
    if re.search(r"NY\s*-?\s*6300\s*\(\s*C\s*\)", text, re.I):
        return "NY6300C"
    match = re.search(r"NY\s*-?\s*(?:P\dC?|[A-Z]?\d{3,4}[A-Z0-9]*P?C?)", text, re.I)
    return _norm_product(match.group(0)) if match else ""


def _shengyi_product_aliases(product: str, *, for_pp: bool) -> set[str]:
    norm = _norm_product(product)
    aliases = {norm}
    if norm in {"NYP1C", "NYP2C", "NYP3C"}:
        aliases.add(norm[:-1])
    if norm in {"NYP1", "NYP2", "NYP3"}:
        aliases.add(f"{norm}C")
    if for_pp:
        if norm.endswith("P"):
            aliases.add(norm[:-1])
        else:
            aliases.add(f"{norm}P")
    return aliases


def _shengyi_extract_thickness_mm(desc: str) -> float | None:
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:±|卤|\+/-|-)\s*\d+(?:\.\d+)?\s*MM", desc, re.I)
    if match:
        return float(match.group(1))
    return _extract_thickness_mm(desc)


def _shengyi_extract_rc(desc: str) -> float | None:
    match = re.search(r"RC\s*[:=：]?\s*(\d+(?:\.\d+)?)", desc, re.I)
    if not match:
        return _extract_rc(desc)
    value = float(match.group(1))
    return value * 100 if value <= 1 else value


def _shengyi_extract_size(desc: str) -> tuple[float | None, float | None]:
    matches = re.findall(r"(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH)?\s*[*xX×]\s*(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH)?", desc, re.I)
    valid: list[tuple[float, float]] = []
    for first, second in matches:
        a, b = float(first), float(second)
        if 0 < a <= 120 and 0 < b <= 120 and not (int(a) in {106, 1067, 1078, 1080, 1506, 2113, 2116, 2313, 3313, 7628} or int(b) in {106, 1067, 1078, 1080, 1506, 2113, 2116, 2313, 3313, 7628}):
            valid.append((a, b))
    if valid:
        return valid[-1]
    return _extract_size(desc, ignore_decimal=False)


def _shengyi_extract_pp_width(desc: str) -> float | None:
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH)?\s*\(\s*纬\s*\)\s*[*xX×]\s*\d+(?:\.\d+)?\s*M", desc, re.I)
    return float(match.group(1)) if match else None


def _shengyi_norm_foil(value: Any) -> str:
    foil = _norm_foil(value)
    aliases = {
        "HS1-M2-VSP": "HVLP1",
        "HS2-M2-VSP": "HVLP2",
        "HVLP": "HVLP1",
    }
    return aliases.get(foil, foil)


def _shengyi_extract_pp_piece(desc: str) -> dict[str, float] | None:
    if re.search(r"[*xX×脳]\s*\d+(?:\.\d+)?\s*M\b", desc, re.I):
        return None
    matches = re.findall(r"(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH)?\s*[*xX×脳]\s*(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH)?", desc, re.I)
    valid: list[tuple[float, float]] = []
    glass_codes = {106, 1067, 1078, 1080, 1506, 2113, 2116, 2313, 3313, 7628}
    for first, second in matches:
        a, b = float(first), float(second)
        if 0 < a <= 120 and 0 < b <= 120 and int(a) not in glass_codes and int(b) not in glass_codes:
            valid.append((a, b))
    if not valid:
        return None
    return _shengyi_piece_axes(*valid[-1], mother_hint=49.5)


def _shengyi_piece_axes(first: float, second: float, mother_hint: float | None = None) -> dict[str, float]:
    radial, latitudinal = first, second
    mother = mother_hint if mother_hint and mother_hint > 0 else latitudinal
    return {"radial": radial, "latitudinal": latitudinal, "mother_width": mother}


def _shengyi_piece_formula_price(radial: float, latitudinal: float, mother_width: float | None, sf_price: float | None) -> dict[str, Any]:
    if sf_price is None:
        return {"ok": False, "reason": "Shengyi small-piece price requires SF/RMB or Per SF"}
    if not mother_width or mother_width <= 0 or latitudinal <= 0 or radial <= 0:
        return {"ok": False, "reason": "Shengyi small-piece dimensions are incomplete"}
    opens = math.floor(mother_width / latitudinal)
    if opens < 1:
        return {"ok": False, "reason": "Shengyi small-piece split count is less than 1"}
    price = _round_money(radial * 0.0254 / opens * float(sf_price))
    formula = f"round({radial:g}*0.0254/{opens}*{float(sf_price):.6g},2)"
    return {"ok": True, "price": price, "label": f"{radial:g}*{latitudinal:g}", "formula": formula, "opens": opens}


def _shengyi_copper_aliases(copper: str) -> set[str]:
    aliases = {copper, _reverse_copper(copper)}
    mapped = {
        "1/H": "1/1",
        "H/1": "1/1",
        "H/2": "1/2",
        "2/H": "2/1",
    }
    if copper in mapped:
        aliases.add(mapped[copper])
        aliases.add(_reverse_copper(mapped[copper]))
    return aliases


def _shengyi_thickness_matches(row: ExtCclRule, thickness_mil: float | None, thickness_mm: float | None) -> bool:
    if thickness_mm is not None and row.thickness_mm is not None and abs(row.thickness_mm - thickness_mm) <= 0.012:
        return True
    if thickness_mil is not None and row.thickness_mil is not None and abs(row.thickness_mil - thickness_mil) <= 0.35:
        return True
    if thickness_mil is not None and row.thickness_mm is not None and abs(row.thickness_mm - thickness_mil * 0.0254) <= 0.012:
        return True
    return False


def _shengyi_thickness_distance(row: ExtCclRule, thickness_mil: float | None, thickness_mm: float | None) -> float:
    distances: list[float] = []
    if thickness_mm is not None and row.thickness_mm is not None:
        distances.append(abs(row.thickness_mm - thickness_mm))
    if thickness_mil is not None and row.thickness_mil is not None:
        distances.append(abs(row.thickness_mil - thickness_mil) * 0.0254)
    if thickness_mil is not None and row.thickness_mm is not None:
        distances.append(abs(row.thickness_mm - thickness_mil * 0.0254))
    return min(distances) if distances else 999.0


def _shengyi_adjusted_sf(row: ExtCclRule, requested_foil: str, copper: str) -> dict[str, Any]:
    sf = row.prices.get("SF")
    if sf is None:
        return {"ok": False, "reason": "生益CCL报价行缺少 Per SF 面积价"}
    row_foil = _shengyi_norm_foil(row.foil) if row.foil else ""
    foil = _shengyi_norm_foil(requested_foil) if requested_foil else row_foil
    if not requested_foil or foil == row_foil:
        return {"ok": True, "sf": float(sf), "note": ""}
    sheet = row.sheet.upper()
    if sheet in {"NY2150", "NY2170", "NY2170H"} and row_foil == "HTE" and foil.startswith("RTF"):
        return {"ok": True, "sf": float(sf) * 1.35, "note": "；RTF铜箔按基板价上调35%"}
    if sheet in {"NY3170M2", "NY6180L", "NY6300S", "NY6300SN"} and row_foil == "RTF2" and foil in {"RTF", "HTE"}:
        return {"ok": True, "sf": float(sf), "note": "；RTF/HTE按RTF2同价"}
    if sheet == "NY6200" and row_foil == "RTF" and foil == "HTE":
        return {"ok": True, "sf": float(sf), "note": "；HTE与RTF同价"}
    if sheet in {"NY3170M", "NY6200"} and row_foil == "RTF" and foil == "RTF2":
        adder = _shengyi_foil_adder(copper, half=0.56, one=0.84)
        return {"ok": True, "sf": float(sf) + adder, "note": f"；RTF2面积加价{adder:.2f}/SF"}
    if sheet == "NY3170LK" and row_foil == "RTF" and foil == "RTF2":
        adder = _shengyi_foil_adder(copper, half=0.6, one=0.9)
        return {"ok": True, "sf": float(sf) + adder, "note": f"；RTF2面积加价{adder:.2f}/SF"}
    if sheet in {"NY6300", "NY6300C"} and row_foil in {"HVLP1", "HVLP2", "HS2", "HS2-M2"} and foil == "RTF2":
        adder = -_shengyi_foil_adder(copper, half=0.75, one=1.13)
        return {"ok": True, "sf": float(sf) + adder, "note": f"；RTF2面积减价{abs(adder):.2f}/SF"}
    if sheet in {"NY6300", "NY6300C"} and row_foil in {"HVLP1", "HVLP2", "HS2", "HS2-M2"} and foil == "RTF3":
        adder = -_shengyi_foil_adder(copper, half=0.37, one=0.565)
        return {"ok": True, "sf": float(sf) + adder, "note": f"；RTF3面积减价{abs(adder):.2f}/SF"}
    if sheet in {"NY6300S", "NY6300SN"} and row_foil == "RTF2" and foil == "RTF3":
        adder = _shengyi_foil_adder(copper, half=0.4, one=0.65)
        return {"ok": True, "sf": float(sf) + adder, "note": f"；RTF3面积加价{adder:.2f}/SF"}
    if sheet in {"NY6300S", "NY6300SN"} and row_foil == "RTF2" and foil == "RTF4":
        adder = _shengyi_foil_adder(copper, half=0.94, one=1.5)
        return {"ok": True, "sf": float(sf) + adder, "note": f"；RTF4面积加价{adder:.2f}/SF"}
    if sheet in {"NY6300S", "NY6300SN"} and row_foil == "RTF2" and foil == "HVLP1":
        adder = _shengyi_foil_adder(copper, half=1.1, one=1.8)
        return {"ok": True, "sf": float(sf) + adder, "note": f"；HVLP1面积加价{adder:.2f}/SF"}
    return {"ok": False, "reason": f"生益{row.sheet}铜箔{foil}需另行报价或未在底部规则明确"}


def _shengyi_foil_adder(copper: str, *, half: float, one: float) -> float:
    total = 0.0
    for part in copper.split("/"):
        if part == "H":
            total += half
        elif part == "1":
            total += one
        else:
            number = _to_float(part)
            total += one * number if number is not None else 0
    return total


def _shengyi_ccl_size_price(row: ExtCclRule, length_in: float, width_in: float, sf_price: float) -> dict[str, Any]:
    dim = _shengyi_sheet_width(length_in, width_in)
    standard_map = [(37, "36*48", 12.0), (41, "40*48", 13.33), (43, "42*48", 14.0)]
    for target, label, factor in standard_map:
        if abs(dim - target) <= 0.8:
            price = _round_money(sf_price * factor)
            return {"ok": True, "price": price, "label": label, "formula": f"round({sf_price:.6g}*{factor:g},2)"}
    narrow = _shengyi_narrow_price(dim, length_in, width_in, sf_price)
    if narrow["ok"]:
        return narrow
    small = _shengyi_small_piece_price(dim, length_in, width_in, sf_price)
    if small["ok"]:
        return small
    return {"ok": False, "reason": f"生益CCL未找到可明确裁切尺寸：{length_in:g}*{width_in:g}"}


def _shengyi_sheet_width(length_in: float, width_in: float) -> float:
    if abs(length_in - 49) <= 1.0 or abs(length_in - 48) <= 1.0:
        return width_in
    if abs(width_in - 49) <= 1.0 or abs(width_in - 48) <= 1.0:
        return length_in
    return max(length_in, width_in)


def _shengyi_narrow_price(dim: float, length_in: float, width_in: float, sf_price: float) -> dict[str, Any]:
    short_side = min(length_in, width_in)
    long_side = max(length_in, width_in)
    if abs(long_side - 43) > 0.8:
        return {"ok": False}
    if abs(short_side - 37) <= 0.8:
        base = _round_money(sf_price * 12)
        price = _round_money(_round_money(base * 43 / 48) * 1.07)
        return {"ok": True, "price": price, "label": "37*43", "formula": f"round(round(round({sf_price:.6g}*12,2)*43/48,2)*1.07,2)"}
    if abs(short_side - 41) <= 0.8:
        base = _round_money(sf_price * 13.33)
        price = _round_money(_round_money(base * 43 / 48) * 1.07)
        return {"ok": True, "price": price, "label": "41*43", "formula": f"round(round(round({sf_price:.6g}*13.33,2)*43/48,2)*1.07,2)"}
    return {"ok": False}


def _shengyi_small_piece_price(dim: float, length_in: float, width_in: float, sf_price: float) -> dict[str, Any]:
    long_side = max(length_in, width_in)
    if not any(abs(long_side - value) <= 1.0 for value in (48.0, 49.0, 49.5)):
        return {"ok": False}
    axes = _shengyi_piece_axes(length_in, width_in, mother_hint=long_side)
    radial = axes["radial"]
    latitudinal = axes["latitudinal"]
    price = _round_money(radial * latitudinal / 144 * sf_price)
    formula = f"round({radial:g}*{latitudinal:g}/144*{sf_price:.6g},2)"
    return {"ok": True, "price": price, "label": f"{radial:g}*{latitudinal:g}", "formula": formula}


def _load_zhongfu_rules(rule_path: str | Path) -> ExtRules:
    wb = load_workbook_compat(rule_path, data_only=True)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    for ws in wb.worksheets:
        max_col = min(ws.max_column, 30)
        for row_idx in range(1, min(ws.max_row, 80) + 1):
            values = [_text(ws.cell(row_idx, col).value).replace("\n", "").replace("\u00a0", " ").strip() for col in range(1, max_col + 1)]
            if _zhongfu_is_ccl_header(values):
                _load_zhongfu_ccl_rows(ws, row_idx, values, ccl_rows)
            if _zhongfu_is_pp_header(values):
                _load_zhongfu_pp_rows(ws, row_idx, values, pp_rows)
    return ExtRules("zhongfu", pp_rows, ccl_rows)


def _load_zhongfu_ccl_rows(ws, header_row: int, headers: list[str], ccl_rows: list[ExtCclRule]) -> None:
    product_col = _find_header_contains(headers, {"型号", "产品型号", "产品类别"})
    mm_col = _find_header_contains(headers, {"公制(mm)", "厚度(mm)", "厚度MM", "mm"}) or 1
    mil_col = _find_header_contains(headers, {"厚度(mil)", "mil"})
    copper_col = _find_header_contains(headers, {"铜厚"}) or 2
    contain_col = _find_header_contains(headers, {"是否含铜"})
    stack_col = _find_header_contains(headers, {"配本", "组合叠构", "叠构"}) or 4
    foil_col = _find_header_contains(headers, {"标准铜箔", "铜箔"})
    price_cols = _zhongfu_ccl_price_columns(headers)
    product = _zhongfu_product_from_text(ws.title)
    blank_streak = 0
    started = False
    for data_row in range(header_row + 1, ws.max_row + 1):
        row_values = [_text(ws.cell(data_row, col).value) for col in range(1, min(ws.max_column, 30) + 1)]
        if not any(row_values):
            if started:
                blank_streak += 1
                if blank_streak >= 50:
                    break
            continue
        started = True
        blank_streak = 0
        if _zhongfu_is_pp_header(row_values):
            break
        row_product = _zhongfu_product_from_text(row_values[product_col - 1] if product_col and product_col <= len(row_values) else "")
        if row_product:
            product = row_product
        thickness_mm = _to_float(row_values[mm_col - 1] if mm_col <= len(row_values) else None)
        thickness_mil = _to_float(row_values[mil_col - 1] if mil_col and mil_col <= len(row_values) else None)
        copper = _norm_copper(row_values[copper_col - 1] if copper_col <= len(row_values) else "")
        contain = _zhongfu_norm_copper_included(row_values[contain_col - 1] if contain_col and contain_col <= len(row_values) else "")
        stack = _norm_stack(row_values[stack_col - 1] if stack_col <= len(row_values) else "")
        foil = _norm_foil(row_values[foil_col - 1] if foil_col and foil_col <= len(row_values) else "")
        prices = {key: _to_float(ws.cell(data_row, col).value) for col, key in price_cols.items()}
        prices = {key: value for key, value in prices.items() if value is not None}
        if product and thickness_mm is not None and copper and stack and prices:
            ccl_rows.append(ExtCclRule(data_row, ws.title, product, thickness_mm, thickness_mil, copper, foil, stack, prices, contain))


def _load_zhongfu_pp_rows(ws, header_row: int, headers: list[str], pp_rows: list[ExtPpRule]) -> None:
    product_col = _find_header_contains(headers, {"中TG型号", "型号", "产品型号"})
    glass_col = _find_header_contains(headers, {"布种", "Glass type"}) or 2
    rc_col = _find_header_contains(headers, {"含胶量", "RC含量从", "Resin Content"}) or 3
    rc_to_col = _find_header_contains(headers, {"RC含量到"})
    width_col = _find_header_contains(headers, {"宽度", "卷状宽度"})
    length_col = _find_header_contains(headers, {"长度", "每卷长度"})
    price_col = _zhongfu_pp_price_col(headers)
    product = _zhongfu_product_from_text(ws.title)
    blank_streak = 0
    started = False
    for data_row in range(header_row + 1, ws.max_row + 1):
        row_values = [_text(ws.cell(data_row, col).value) for col in range(1, min(ws.max_column, 30) + 1)]
        if not any(row_values):
            if started:
                blank_streak += 1
                if blank_streak >= 50:
                    break
            continue
        started = True
        blank_streak = 0
        row_product = _zhongfu_product_from_text(row_values[product_col - 1] if product_col and product_col <= len(row_values) else "")
        if row_product:
            product = row_product
        glass = _zhongfu_norm_glass(row_values[glass_col - 1] if glass_col <= len(row_values) else "")
        rc_min, rc_max = _parse_rc_range(row_values[rc_col - 1] if rc_col <= len(row_values) else "")
        if rc_to_col and rc_to_col <= len(row_values):
            rc_to = _to_float(row_values[rc_to_col - 1])
            if rc_to is not None:
                rc_max = rc_to * 100 if rc_to <= 1 else rc_to
        width = _to_float(row_values[width_col - 1] if width_col and width_col <= len(row_values) else None)
        length = _length_int(row_values[length_col - 1] if length_col and length_col <= len(row_values) else None)
        price = _to_float(ws.cell(data_row, price_col).value) if price_col else None
        if product and glass and rc_min is not None and rc_max is not None and price is not None:
            pp_rows.append(ExtPpRule(data_row, ws.title, product, glass, rc_min, rc_max, length, width, price))


def _calculate_zhongfu_spec(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    if _zhongfu_looks_like_pp(desc):
        return _calculate_zhongfu_pp(desc, rules)
    return _calculate_zhongfu_ccl(desc, rules, quantity=quantity)


def _calculate_zhongfu_pp(desc: str, rules: ExtRules) -> ExtCalcResult:
    if _zhongfu_is_pp_small_piece(desc):
        return ExtCalcResult("失败", "PP", "待确认", "", "", "", "中富PP小片不计算")
    product = _extract_product(desc) or _zhongfu_product_from_text(desc)
    glass = _zhongfu_norm_glass(_extract_glass(desc))
    rc = _extract_rc(desc)
    if not product or not glass or rc is None:
        return ExtCalcResult("失败", "PP", "待确认", "", "", "", "中富PP规格缺少型号、玻布或RC")
    products = _zhongfu_pp_product_aliases(product)
    matches = [
        row
        for row in rules.pp_rows
        if row.product in products
        and row.glass == glass
        and row.rc_min is not None
        and row.rc_max is not None
        and row.rc_min - 0.001 <= rc <= row.rc_max + 0.001
    ]
    if not matches:
        return ExtCalcResult("失败", "PP", "待确认", "", "", "", "未命中中富PP报价：型号、玻布或RC不匹配")
    product_norm = _norm_product(product)
    best = sorted(matches, key=lambda row: (0 if row.product == product_norm else 1, row.excel_row))[0]
    price = _round_money(best.price)
    note = f"命中中富PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，型号={best.product}，玻布={glass}，RC={rc:g}，每米价={price:.2f}"
    return ExtCalcResult("成功", "PP", price, "", _fmt_width(best.width), str(best.length or ""), note, best.excel_row, best.sheet)


def _calculate_zhongfu_ccl(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    product = _extract_product(desc) or _zhongfu_product_from_text(desc)
    thickness_mm = _zhongfu_extract_thickness_mm(desc)
    copper = _extract_copper(desc)
    contain = _zhongfu_extract_copper_included(desc)
    stack = _extract_stack(desc)
    length_in, width_in = _extract_size(desc, ignore_decimal=False)
    if not product or thickness_mm is None or not copper or length_in is None or width_in is None:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "中富CCL规格缺少型号、厚度、铜厚或尺寸")
    product_norm = _norm_product(product)
    rows = [
        row
        for row in rules.ccl_rows
        if row.product == product_norm
        and _thickness_matches(row, None, thickness_mm)
        and row.copper == copper
        and (not row.kind or not contain or row.kind == contain)
    ]
    if stack:
        rows = [row for row in rows if row.stack == stack]
    elif len(rows) != 1:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "中富CCL规格缺少配本，无法唯一确定报价行")
    if not rows:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "未命中中富CCL报价：型号、厚度、铜厚、是否含铜或配本不匹配")
    if len(rows) > 1:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "中富CCL命中多条报价行，需确认配本")
    row = rows[0]
    price_result = _zhongfu_ccl_size_price(row, length_in, width_in)
    if not price_result["ok"]:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", price_result["reason"])
    price = _round_money(price_result["price"])
    total = _calc_total(quantity, price)
    note = (
        f"命中中富CCL报价 Sheet {row.sheet} 第 {row.excel_row} 行，"
        f"型号={row.product}，厚度={row.thickness_mm:g}mm，铜厚={row.copper}，配本={row.stack}，"
        f"尺寸列={price_result['label']}，公式={price_result['formula']}"
    )
    return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, price_result["label"])


def _zhongfu_ccl_size_price(row: ExtCclRule, length_in: float, width_in: float) -> dict[str, Any]:
    direct_key, direct_label = _zhongfu_direct_price_key(length_in, width_in, row.prices)
    if direct_key and row.prices.get(direct_key) is not None:
        price = float(row.prices[direct_key])
        return {"ok": True, "price": price, "label": direct_label, "formula": f"{price:.6g}"}
    cut_result = _zhongfu_ccl_cut_price(row, length_in, width_in)
    if cut_result["ok"]:
        return cut_result
    return {"ok": False, "reason": f"中富CCL未找到可裁切母板：{length_in:g}*{width_in:g}"}


def _zhongfu_ccl_cut_price(row: ExtCclRule, length_in: float, width_in: float) -> dict[str, Any]:
    candidates = [
        (74, 49, "37", 2),
        (82, 49, "41", 2),
        (86, 49, "43", 2),
        (37, 49, "37", 1),
        (41, 49, "41", 1),
        (43, 49, "43", 1),
    ]
    valid: list[dict[str, Any]] = []
    for parent_w, parent_h, source_key, factor in candidates:
        parent_price = row.prices.get(source_key)
        if parent_price is None:
            continue
        for piece_w, piece_h in ((length_in, width_in), (width_in, length_in)):
            if piece_w <= 0 or piece_h <= 0:
                continue
            opens_w = math.floor((parent_w + 1e-9) / piece_w)
            opens_h = math.floor((parent_h + 1e-9) / piece_h)
            opens = opens_w * opens_h
            if opens <= 1:
                continue
            fit_error = abs(piece_w * opens_w - parent_w) + abs(piece_h * opens_h - parent_h)
            if fit_error <= 0.8:
                valid.append(
                    {
                        "price": float(parent_price) * factor / opens,
                        "parent": f"{parent_w}*{parent_h}",
                        "factor": factor,
                        "opens": opens,
                        "fit_error": fit_error,
                        "parent_price": float(parent_price),
                    }
                )
    if not valid:
        return {"ok": False, "reason": f"中富CCL未找到可裁切母板：{length_in:g}*{width_in:g}"}
    best = sorted(valid, key=lambda item: (item["fit_error"], -item["opens"], item["parent"]))[0]
    formula = f"{best['parent_price']:.6g}*{best['factor']:g}/{best['opens']:g}"
    label = f"{best['parent']}/{best['opens']:g}"
    return {"ok": True, "price": best["price"], "label": label, "formula": formula}


def _zhongfu_is_ccl_header(values: list[str]) -> bool:
    compact = {_text(value).replace(" ", "") for value in values}
    has_thickness = any("公制(mm)" in value or "厚度(mm)" in value or value == "mm" for value in compact)
    return has_thickness and "铜厚" in compact and "配本" in compact and any("37*49" in value for value in compact)


def _zhongfu_is_pp_header(values: list[str]) -> bool:
    compact = {_text(value).replace(" ", "") for value in values}
    has_rc = any("含胶量" in value or "RC含量从" in value or "ResinContent" in value for value in compact)
    has_price = any("价格/米" in value or "报价" in value for value in compact)
    return "布种" in compact and has_rc and has_price


def _zhongfu_ccl_price_columns(headers: list[str]) -> dict[int, str]:
    price_cols: dict[int, str] = {}
    for idx, header in enumerate(headers, start=1):
        normalized = _text(header).replace(" ", "")
        if "37*49" in normalized:
            price_cols[idx] = "37"
        elif "41*49" in normalized:
            price_cols[idx] = "41"
        elif "43*49" in normalized:
            price_cols[idx] = "43"
    return price_cols


def _zhongfu_pp_price_col(headers: list[str]) -> int | None:
    for idx, header in enumerate(headers, start=1):
        normalized = _text(header).replace(" ", "")
        if "ROLL" in normalized.upper():
            continue
        if "价格/米" in normalized or "报价" in normalized or "PerM" in normalized:
            return idx
    return None


def _zhongfu_product_from_text(value: Any) -> str:
    text = _text(value).upper().replace(" ", "").replace("-", "")
    match = re.search(r"NY(?:2150HP|2170HP|2150H|2170H|3150HF|3176HF|6300SL|6300S|6200P?)", text)
    if match:
        return _norm_product(match.group(0))
    match = re.search(r"NY[A-Z]?\d{3,4}[A-Z]{0,4}P?", text)
    return _norm_product(match.group(0)) if match else ""


def _zhongfu_pp_product_aliases(product: str) -> set[str]:
    product_norm = _norm_product(product)
    aliases = {product_norm}
    if product_norm.endswith("P"):
        aliases.add(product_norm[:-1])
    else:
        aliases.add(f"{product_norm}P")
    return aliases


def _zhongfu_norm_glass(value: Any) -> str:
    text = _text(value).upper().replace("A", "")
    return _norm_glass(text)


def _zhongfu_norm_copper_included(value: Any) -> str:
    text = _text(value)
    if "不含" in text or text == "否":
        return "不含铜"
    if "含" in text or text == "是":
        return "含铜"
    return ""


def _zhongfu_extract_copper_included(desc: str) -> str:
    if "不含铜" in desc:
        return "不含铜"
    if "含铜" in desc:
        return "含铜"
    return ""


def _zhongfu_extract_thickness_mm(desc: str) -> float | None:
    product_pattern = r"NY\s*-?\s*(?:2150HP|2170HP|2150H|2170H|3150HF|3176HF|6300SL|6300S|6200P?)"
    match = re.search(product_pattern + r".{0,12}?(\d+(?:\.\d+)?)\s*(?:MM|mm)?\s*[-_]", desc, re.I)
    if match:
        return float(match.group(1))
    match = re.search(r"(\d+(?:\.\d+)?)\s*MM\b", desc, re.I)
    if match:
        return float(match.group(1))
    return _extract_thickness_mm(desc)


def _zhongfu_direct_price_key(length_in: float, width_in: float, prices: dict[str, float | None]) -> tuple[str, str]:
    dim = _zhongfu_sheet_width(length_in, width_in)
    for target, key in [(37, "37"), (41, "41"), (43, "43")]:
        if abs(dim - target) <= 0.8 and key in prices:
            return key, f"{key}*49"
    return "", ""


def _zhongfu_sheet_width(length_in: float, width_in: float) -> float:
    if abs(length_in - 49) <= 0.8:
        return width_in
    if abs(width_in - 49) <= 0.8:
        return length_in
    return max(length_in, width_in)


def _zhongfu_looks_like_pp(desc: str) -> bool:
    upper = desc.upper()
    return "RC" in upper and (_extract_glass(desc) != "" or re.search(r"P\d{3,4}\s*RC", upper) is not None)


def _zhongfu_is_pp_small_piece(desc: str) -> bool:
    if "经" in desc or "經" in desc:
        return True
    return re.search(r"\d+(?:\.\d+)?\s*MM\s*[*xX×]\s*\d+(?:\.\d+)?\s*MM", desc, re.I) is not None


def _load_guanghe_rules(rule_path: str | Path) -> ExtRules:
    wb = load_workbook_compat(rule_path, data_only=True)
    pp_rows: list[ExtPpRule] = []
    ccl_rows: list[ExtCclRule] = []
    for ws in wb.worksheets:
        max_col = min(ws.max_column, 30)
        for row_idx in range(1, ws.max_row + 1):
            values = [_text(ws.cell(row_idx, col).value) for col in range(1, max_col + 1)]
            normalized = [value.replace("\n", "").replace("\u00a0", " ").strip() for value in values]
            if _guanghe_is_ccl_header(normalized):
                _load_guanghe_ccl_rows(ws, row_idx, normalized, ccl_rows)
            if _guanghe_is_pp_header(normalized):
                _load_guanghe_pp_rows(ws, row_idx, normalized, pp_rows)
    return ExtRules("guanghe", pp_rows, ccl_rows)


def _load_guanghe_ccl_rows(ws, header_row: int, headers: list[str], ccl_rows: list[ExtCclRule]) -> None:
    product_col = _find_header_contains(headers, {"产品类别", "Products"}) or 1
    mm_col = _find_header_contains(headers, {"厚度mm", "厚度 MM", "mm"}) or 2
    mil_col = _find_header_contains(headers, {"厚度mil", "mil"}) or 3
    copper_col = _find_header_contains(headers, {"铜厚"}) or 4
    foil_col = _find_header_contains(headers, {"铜箔类型"}) or 5
    stack_col = _find_header_contains(headers, {"组合叠构", "叠构"}) or 6
    price_cols = _guanghe_ccl_price_columns(headers)
    product = _guanghe_product_from_text(ws.title)
    thickness_mm: float | None = None
    thickness_mil: float | None = None
    foil = ""
    stack = ""
    for data_row in range(header_row + 1, ws.max_row + 1):
        row_values = [_text(ws.cell(data_row, col).value) for col in range(1, 31)]
        if _guanghe_is_pp_header(row_values) or any("产品型号" in value for value in row_values):
            break
        if not any(row_values):
            continue
        row_product = _guanghe_product_from_text(row_values[product_col - 1] if product_col <= len(row_values) else "")
        if row_product:
            product = row_product
        row_thickness_mm = _to_float(row_values[mm_col - 1] if mm_col <= len(row_values) else None)
        row_thickness_mil = _to_float(row_values[mil_col - 1] if mil_col <= len(row_values) else None)
        if row_thickness_mm is not None:
            thickness_mm = row_thickness_mm
        if row_thickness_mil is not None:
            thickness_mil = row_thickness_mil
        copper = _norm_copper(row_values[copper_col - 1] if copper_col <= len(row_values) else "")
        row_foil = _norm_foil(row_values[foil_col - 1] if foil_col <= len(row_values) else "")
        row_stack = _norm_stack(row_values[stack_col - 1] if stack_col <= len(row_values) else "")
        if row_foil:
            foil = row_foil
        if row_stack:
            stack = row_stack
        prices = {key: _to_float(ws.cell(data_row, col).value) for col, key in price_cols.items()}
        prices = {key: value for key, value in prices.items() if value is not None}
        if not product or not copper or not stack or not prices:
            continue
        ccl_rows.append(ExtCclRule(data_row, ws.title, product, thickness_mm, thickness_mil, copper, foil, stack, prices))


def _load_guanghe_pp_rows(ws, header_row: int, headers: list[str], pp_rows: list[ExtPpRule]) -> None:
    product_col = _find_header_contains(headers, {"Products", "产品类别"}) or 1
    glass_col = _find_header_contains(headers, {"Glass type", "产品型号"}) or 2
    rc_col = _find_header_contains(headers, {"Resin Content", "树脂含量"}) or 3
    length_col = _find_header_contains(headers, {"Length", "标准卷装长度"}) or 4
    width_col = _find_header_contains(headers, {"Width", "标准卷装宽度"}) or 5
    per_m_col = _guanghe_first_header_col(headers, {"M未税", "Per M"})
    per_43_col = _guanghe_43_price_col(headers)
    product = _guanghe_product_from_text(ws.title)
    for data_row in range(header_row + 1, ws.max_row + 1):
        row_values = [_text(ws.cell(data_row, col).value) for col in range(1, 31)]
        if not any(row_values):
            continue
        row_product = _guanghe_product_from_text(row_values[product_col - 1] if product_col <= len(row_values) else "")
        if row_product:
            product = row_product
        glass = _norm_glass(row_values[glass_col - 1] if glass_col <= len(row_values) else "")
        rc_min, rc_max = _parse_rc_range(row_values[rc_col - 1] if rc_col <= len(row_values) else "")
        length = _length_int(row_values[length_col - 1] if length_col <= len(row_values) else None)
        width = _to_float(row_values[width_col - 1] if width_col <= len(row_values) else None)
        if width is not None and 0 < width <= 1:
            width *= 100
        price = _to_float(ws.cell(data_row, per_m_col).value) if per_m_col else None
        price_43 = _to_float(ws.cell(data_row, per_43_col).value) if per_43_col else None
        if product and glass and rc_min is not None and rc_max is not None and (price is not None or price_43 is not None):
            pp_rows.append(ExtPpRule(data_row, ws.title, product, glass, rc_min, rc_max, length, width, price, price_43))


def _calculate_guanghe_spec(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    if _looks_like_pp(desc) or desc.upper().startswith("PP"):
        return _calculate_guanghe_pp(desc, rules)
    return _calculate_guanghe_ccl(desc, rules, quantity=quantity)


def _calculate_guanghe_pp(desc: str, rules: ExtRules) -> ExtCalcResult:
    product = _extract_product(desc) or _guanghe_product_from_text(desc)
    glass = _extract_glass(desc)
    rc = _extract_rc(desc)
    width = _guanghe_extract_pp_width(desc) or _extract_width(desc)
    if not product or not glass or rc is None:
        return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width), "", "广合PP规格缺少型号、玻布或RC")
    products = _guanghe_product_aliases(product)
    matches = [
        row
        for row in rules.pp_rows
        if row.product in products
        and row.glass == glass
        and row.rc_min is not None
        and row.rc_max is not None
        and row.rc_min - 0.001 <= rc <= row.rc_max + 0.001
    ]
    if not matches:
        return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width), "", "未命中广合PP报价：型号、玻布或RC不匹配")
    product_norm = _norm_product(product)
    best = sorted(
        matches,
        key=lambda row: (
            0 if row.product == product_norm else 1,
            row.excel_row,
        ),
    )[0]
    small_piece = _guanghe_extract_pp_small_piece(desc)
    if small_piece:
        cut_width_in, full_width_in = small_piece
        split = math.floor(full_width_in / cut_width_in) if cut_width_in else 0
        if split <= 0:
            return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(cut_width_in), "", f"广合PP小片纬向无法一开：{cut_width_in:g}inch")
        if best.price is None:
            return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(cut_width_in), "", "命中广合PP报价行但每米价为空")
        raw_price = cut_width_in * 0.0254 / split * float(best.price)
        price = _round_money(raw_price)
        note = (
            f"命中广合PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，"
            f"型号={best.product}，玻布={glass}，RC={rc:g}，小片={cut_width_in:g}*{full_width_in:g}inch，"
            f"公式={cut_width_in:g}*0.0254/{split}*{float(best.price):.6g}"
        )
        return ExtCalcResult("成功", "PP", price, "", _fmt_width(cut_width_in), "", note, best.excel_row, best.sheet)
    use_43 = width is not None and width <= 43.8
    raw_price = best.roll_price if use_43 else best.price
    if raw_price is None:
        missing_label = "43inch每米价" if use_43 else "每米价"
        return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width), "", f"命中广合PP报价行但{missing_label}为空")
    price = _round_money(raw_price)
    price_label = "43inch每米价" if use_43 else "每米价"
    note = (
        f"命中广合PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，"
        f"型号={best.product}，玻布={glass}，RC={rc:g}，{price_label}={price:.2f}"
    )
    return ExtCalcResult("成功", "PP", price, "", _fmt_width(width or best.width), "", note, best.excel_row, best.sheet)


def _calculate_guanghe_ccl(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    product = _extract_product(desc) or _guanghe_product_from_text(desc)
    thickness_mil = _guanghe_extract_thickness_mil(desc)
    thickness_mm = thickness_mil * 0.0254 if thickness_mil is not None else _extract_thickness_mm(desc)
    copper = _extract_copper(desc)
    foil = _extract_foil(desc)
    stack = _extract_stack(desc)
    length_in, width_in = _guanghe_extract_size(desc)
    if not product or thickness_mm is None or not copper or not stack or length_in is None or width_in is None:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "广合CCL规格缺少型号、厚度、铜厚、叠构或尺寸")
    products = _guanghe_product_aliases(product)
    product_rows = [
        row
        for row in rules.ccl_rows
        if row.product in products
        and _thickness_matches(row, thickness_mil, thickness_mm)
    ]
    copper_rows = [row for row in product_rows if row.copper == copper or row.copper == _reverse_copper(copper)]
    stack_rows = [row for row in copper_rows if row.stack == stack]
    candidates = [row for row in stack_rows if not row.foil or not foil or row.foil == foil]
    if not candidates:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "未命中广合CCL报价：型号、厚度、铜厚、叠构或铜箔不匹配")
    last_reason = ""
    for row in sorted(candidates, key=lambda item: (0 if item.product == _norm_product(product) else 1, item.excel_row)):
        price_result = _guanghe_ccl_size_price(row, length_in, width_in)
        if price_result["ok"]:
            price = _round_money(price_result["price"])
            total = _calc_total(quantity, price)
            thickness_text = f"{row.thickness_mil:g}mil" if row.thickness_mil is not None else ""
            note = (
                f"命中广合CCL报价 Sheet {row.sheet} 第 {row.excel_row} 行，"
                f"型号={row.product}，厚度={thickness_text}，铜厚={row.copper}，"
                f"叠构={row.stack}，尺寸列={price_result['label']}，公式={price_result['formula']}"
            )
            return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, price_result["label"])
        last_reason = price_result["reason"]
    return ExtCalcResult("失败", "CCL", "待确认", "", "", "", last_reason or "广合CCL报价行找到，但尺寸未匹配")


def _guanghe_ccl_size_price(row: ExtCclRule, length_in: float, width_in: float) -> dict[str, Any]:
    direct_key, direct_label = _guanghe_direct_price_key(length_in, width_in, row.prices)
    if direct_key and row.prices.get(direct_key) is not None:
        price = float(row.prices[direct_key])
        return {"ok": True, "price": price, "label": direct_label, "formula": f"{price:.6g}"}
    dim = _guanghe_sheet_width(length_in, width_in)
    for target, source_key, factor in [(74, "37", 2), (82, "41", 2), (86, "43", 2)]:
        if abs(dim - target) <= 0.8 and row.prices.get(source_key) is not None:
            base = float(row.prices[source_key])
            return {"ok": True, "price": base * factor, "label": f"{target}*49", "formula": f"{base:.6g}*{factor:g}"}
    cut_result = _guanghe_ccl_cut_price(row, length_in, width_in)
    if cut_result["ok"]:
        return cut_result
    sf_price = row.prices.get("SF")
    if sf_price is not None:
        price = float(sf_price) * length_in * width_in / 144
        formula = f"{float(sf_price):.6g}*{length_in:.6g}*{width_in:.6g}/144"
        return {"ok": True, "price": price, "label": "SF", "formula": formula}
    return {"ok": False, "reason": f"广合CCL未找到可用尺寸报价：{length_in:g}*{width_in:g}"}


def _guanghe_is_ccl_header(values: list[str]) -> bool:
    compact = {_text(value).replace(" ", "") for value in values}
    return "厚度mm" in compact and "铜厚" in compact and ("37*49" in compact or "41*49" in compact or "43*49" in compact)


def _guanghe_is_pp_header(values: list[str]) -> bool:
    compact = {_text(value).replace(" ", "") for value in values}
    return "Glasstype" in compact and "ResinContent" in compact


def _guanghe_product_from_text(value: Any) -> str:
    match = re.search(r"NY\s*-?\s*(?:\d{3,4}[A-Z0-9]*P?|P\dP?)", _text(value), re.I)
    return _norm_product(match.group(0)) if match else ""


def _guanghe_product_aliases(product: str) -> set[str]:
    product_norm = _norm_product(product)
    aliases = {product_norm}
    if product_norm.endswith("P"):
        aliases.add(product_norm[:-1])
    else:
        aliases.add(f"{product_norm}P")
    return aliases


def _guanghe_extract_thickness_mil(desc: str) -> float | None:
    quote_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH|英寸)", desc, re.I)
    if quote_match:
        return float(quote_match.group(1)) * 1000
    return _extract_thickness_mil(desc, _extract_product(desc))


def _guanghe_extract_pp_width(desc: str) -> float | None:
    match = re.search(r"纬向\s*(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH|英寸)?", desc, re.I)
    return float(match.group(1)) if match else None


def _guanghe_extract_pp_small_piece(desc: str) -> tuple[float, float] | None:
    match = re.search(
        r"(?<!RC)(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH|英寸)?\s*[*xX×]\s*(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH|英寸)",
        desc,
        re.I,
    )
    if not match:
        return None
    first = float(match.group(1))
    second = float(match.group(2))
    if first <= 0 or second <= 0 or first > 60 or second > 60:
        return None
    full_width = max(first, second)
    cut_width = min(first, second)
    if full_width < 49:
        return None
    return cut_width, full_width


def _guanghe_extract_size(desc: str) -> tuple[float | None, float | None]:
    match = re.search(
        r"经向\s*(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH|英寸)?\s*[*xX×]?\s*纬向\s*(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH|英寸)?",
        desc,
        re.I,
    )
    if match:
        return float(match.group(1)), float(match.group(2))
    return _extract_size(desc, ignore_decimal=False)


def _guanghe_direct_price_key(length_in: float, width_in: float, prices: dict[str, float | None]) -> tuple[str, str]:
    dim = _guanghe_sheet_width(length_in, width_in)
    for target, key in [(37, "37"), (41, "41"), (43, "43")]:
        if abs(dim - target) <= 0.8 and key in prices:
            return key, f"{key}*49"
    return _direct_price_key(length_in, width_in, prices)


def _guanghe_sheet_width(length_in: float, width_in: float) -> float:
    if abs(length_in - 49) <= 1.0:
        return width_in
    if abs(width_in - 49) <= 1.0:
        return length_in
    return max(length_in, width_in)


def _guanghe_ccl_cut_price(row: ExtCclRule, length_in: float, width_in: float) -> dict[str, Any]:
    candidates = [
        (74, 49, "37", 2),
        (82, 49, "41", 2),
        (86, 49, "43", 2),
        (37, 49, "37", 1),
        (41, 49, "41", 1),
        (43, 49, "43", 1),
    ]
    valid: list[dict[str, Any]] = []
    for parent_w, parent_h, source_key, factor in candidates:
        parent_price = row.prices.get(source_key)
        if parent_price is None:
            continue
        for piece_w, piece_h in ((length_in, width_in), (width_in, length_in)):
            if piece_w <= 0 or piece_h <= 0:
                continue
            opens_w = math.floor((parent_w + 1e-9) / piece_w)
            opens_h = math.floor((parent_h + 1e-9) / piece_h)
            opens = opens_w * opens_h
            if opens <= 1:
                continue
            fit_error = abs(piece_w * opens_w - parent_w) + abs(piece_h * opens_h - parent_h)
            if fit_error <= 0.8:
                valid.append(
                    {
                        "price": float(parent_price) * factor / opens,
                        "parent": f"{parent_w}*{parent_h}",
                        "source_key": source_key,
                        "factor": factor,
                        "opens": opens,
                        "fit_error": fit_error,
                        "parent_price": float(parent_price),
                    }
                )
    if not valid:
        return {"ok": False, "reason": f"广合CCL未找到可裁切母板：{length_in:g}*{width_in:g}"}
    best = sorted(valid, key=lambda item: (item["fit_error"], -item["opens"], item["parent"]))[0]
    formula = f"{best['parent_price']:.6g}*{best['factor']:g}/{best['opens']:g}"
    label = f"{best['parent']}/{best['opens']:g}"
    return {"ok": True, "price": best["price"], "label": label, "formula": formula}


def _guanghe_ccl_price_columns(headers: list[str]) -> dict[int, str]:
    price_cols: dict[int, str] = {}
    for idx, header in enumerate(headers, start=1):
        normalized = _text(header).replace(" ", "")
        if not normalized:
            continue
        if "SF" in normalized:
            price_cols[idx] = "SF"
        elif "37*49" in normalized:
            price_cols[idx] = "37"
        elif "41*49" in normalized:
            price_cols[idx] = "41"
        elif "43*49" in normalized:
            price_cols[idx] = "43"
    return price_cols


def _guanghe_first_header_col(headers: list[str], names: set[str]) -> int | None:
    for idx, header in enumerate(headers, start=1):
        normalized = _text(header).replace(" ", "").replace("\u00a0", "")
        if any(name.replace(" ", "") in normalized for name in names):
            return idx
    return None


def _guanghe_43_price_col(headers: list[str]) -> int | None:
    for idx, header in enumerate(headers, start=1):
        normalized = _text(header).replace(" ", "").replace("\u00a0", "").lower()
        if "43" in normalized and ("m未税" in normalized or "perm" in normalized or "43in" in normalized):
            return idx
    return None


def _row_value(values: list[Any], one_based_col: int) -> Any:
    index = one_based_col - 1
    return values[index] if 0 <= index < len(values) else None


def _workbook_value_sheets(path: str | Path) -> list[tuple[str, list[list[Any]]]]:
    try:
        wb = load_workbook_compat(path, data_only=True)
        return [
            (
                ws.title,
                [[ws.cell(row=row_idx, column=col_idx).value for col_idx in range(1, ws.max_column + 1)] for row_idx in range(1, ws.max_row + 1)],
            )
            for ws in wb.worksheets
        ]
    except Exception:
        return _xlsx_value_sheets_from_xml(path)


_XLSX_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_XLSX_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_XLSX_PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _xlsx_value_sheets_from_xml(path: str | Path) -> list[tuple[str, list[list[Any]]]]:
    with zipfile.ZipFile(Path(path)) as archive:
        shared = _xlsx_shared_strings(archive)
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rels = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root.findall(f"{_XLSX_PKG_REL_NS}Relationship")}
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets: list[tuple[str, list[list[Any]]]] = []
        for sheet in workbook_root.find(f"{_XLSX_MAIN_NS}sheets").findall(f"{_XLSX_MAIN_NS}sheet"):
            name = sheet.attrib["name"]
            rel_id = sheet.attrib.get(f"{_XLSX_REL_NS}id")
            target = rels.get(rel_id or "", "")
            sheet_path = f"xl/{target.lstrip('/')}" if not target.startswith("xl/") else target
            sheet_root = ET.fromstring(archive.read(sheet_path))
            rows: list[list[Any]] = []
            for row in sheet_root.findall(f".//{_XLSX_MAIN_NS}sheetData/{_XLSX_MAIN_NS}row"):
                row_idx = int(row.attrib.get("r", "0"))
                while len(rows) < row_idx:
                    rows.append([])
                values = rows[row_idx - 1]
                for cell in row.findall(f"{_XLSX_MAIN_NS}c"):
                    col_idx = _xlsx_col_index(cell.attrib.get("r", ""))
                    if not col_idx:
                        continue
                    while len(values) < col_idx:
                        values.append(None)
                    values[col_idx - 1] = _xlsx_cell_value(cell, shared)
            sheets.append((name, rows))
        return sheets


def _xlsx_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(node.text or "" for node in item.iter(f"{_XLSX_MAIN_NS}t")) for item in root.findall(f"{_XLSX_MAIN_NS}si")]


def _xlsx_cell_value(cell, shared: list[str]) -> Any:
    value_node = cell.find(f"{_XLSX_MAIN_NS}v")
    raw = value_node.text if value_node is not None else ""
    cell_type = cell.attrib.get("t")
    if cell_type == "s" and raw:
        try:
            return shared[int(raw)]
        except (ValueError, IndexError):
            return raw
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{_XLSX_MAIN_NS}t"))
    if raw and re.fullmatch(r"-?\d+(?:\.\d+)?", raw):
        number = float(raw)
        return int(number) if number.is_integer() else number
    return raw


def _xlsx_col_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return 0
    value = 0
    for char in match.group(1):
        value = value * 26 + ord(char) - 64
    return value


def _load_aoshikang_pp_sheet(ws, pp_rows: list[ExtPpRule]) -> None:
    current_glass = ""
    for row_idx in range(1, ws.max_row):
        label = _text(ws.cell(row_idx, 1).value)
        next_label = _text(ws.cell(row_idx + 1, 1).value)
        segment_products = _aoshikang_pp_segment_aliases(label)
        if not label or not segment_products:
            continue
        if next_label == "布种":
            header_row = row_idx + 1
            special_col_products = _aoshikang_pp_special_products(ws.cell(header_row, 5).value)
            data_row = header_row + 1
        elif _to_float(ws.cell(row_idx + 1, 1).value) is not None and _to_float(ws.cell(row_idx + 1, 4).value) is not None:
            special_col_products = set()
            data_row = row_idx + 1
        else:
            continue
        current_glass = ""
        while data_row <= ws.max_row:
            first = _text(ws.cell(data_row, 1).value)
            if first and _aoshikang_pp_segment_aliases(first) and _to_float(ws.cell(data_row, 4).value) is None:
                break
            if first and data_row < ws.max_row and _text(ws.cell(data_row + 1, 1).value) == "布种":
                break
            if first:
                current_glass = _aoshikang_norm_glass(first)
            rc_min, rc_max = _parse_rc_range(ws.cell(data_row, 2).value)
            normal_price = _to_float(ws.cell(data_row, 4).value)
            special_price = _to_float(ws.cell(data_row, 5).value)
            if current_glass and rc_min is not None:
                if normal_price is not None:
                    for product in segment_products:
                        pp_rows.append(ExtPpRule(data_row, ws.title, product, current_glass, rc_min, rc_max, None, None, normal_price))
                if special_price is not None:
                    for product in special_col_products:
                        pp_rows.append(ExtPpRule(data_row, ws.title, product, current_glass, rc_min, rc_max, None, None, special_price))
            data_row += 1


def _calculate_mingyang_spec(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    if _looks_like_pp(desc) or "PP" in desc.upper():
        return _calculate_mingyang_pp(desc, rules)
    return _calculate_mingyang_ccl(desc, rules, quantity=quantity)


def _calculate_mingyang_pp(desc: str, rules: ExtRules) -> ExtCalcResult:
    product = _extract_product(desc)
    glass = _extract_glass(desc)
    rc = _extract_rc(desc)
    length = _extract_length(desc)
    small_length_m, small_width_mm = _extract_mingyang_pp_small_piece_mm(desc)
    if not product or not glass or rc is None:
        return ExtCalcResult("失败", "PP", "待确认", "", "", _fmt_length(length), "PP规格缺少型号、布种或RC")
    product_norm = _norm_product(product)
    base_product = product_norm[:-1] if product_norm.endswith("P") else product_norm
    products = {product_norm, base_product, f"{base_product}P"}
    matches = [
        row
        for row in rules.pp_rows
        if row.product in products
        and row.glass == glass
        and row.rc_min is not None
        and row.rc_max is not None
        and row.rc_min - 0.001 <= rc <= row.rc_max + 0.001
    ]
    strict_matches = [row for row in matches if length is None or row.length is None or row.length == length]
    candidates = strict_matches or matches
    if not candidates:
        return ExtCalcResult("失败", "PP", "待确认", "", "", _fmt_length(length), "未命中明阳PP报价：型号、布种、RC不匹配")
    best = sorted(
        candidates,
        key=lambda row: (
            0 if row.product == product_norm else 1,
            0 if length is not None and row.length == length else 1,
            row.excel_row,
        ),
    )[0]
    if best.price is None:
        return ExtCalcResult("失败", "PP", "待确认", "", "", _fmt_length(length), "命中PP报价行但单价为空")
    length_note = "" if length is None or best.length in {None, length} else f"，报价卷长{best.length}m与规格{length}m不一致，按规格米数计算"
    if small_length_m is not None and small_width_mm is not None:
        split = math.floor(((best.width or 49.5) * 25.4 + 1e-9) / small_width_mm)
        if split <= 0:
            return ExtCalcResult("失败", "PP", "待确认", "", "", "", f"PP小片纬向无法一开：报价宽幅{best.width or 49.5:g}inch，纬向{small_width_mm:g}mm")
        raw_price = float(best.price) * small_length_m / split
        price = _round_money(raw_price)
        note = (
            f"命中明阳PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，单价={best.price:.6g}，"
            f"径向={small_length_m * 1000:.0f}mm，纬向={small_width_mm:.0f}mm，"
            f"纬向一开{split}，公式={small_length_m:.3f}*{best.price:.6g}/{split}={price:.2f}{length_note}"
        )
        return ExtCalcResult("成功", "PP", price, "", _fmt_width(best.width), "", note, best.excel_row, best.sheet)
    if length is None:
        return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(best.width), "", "PP卷料缺少米数，无法按单价×米数计算")
    price = _round_money(float(best.price) * length)
    note = f"命中明阳PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，PP卷料: 单价{best.price:.6g}×{length}={price:.2f}{length_note}"
    return ExtCalcResult("成功", "PP", price, "", _fmt_width(best.width), _fmt_length(length), note, best.excel_row, best.sheet)


def _calculate_mingyang_ccl(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    product = _extract_product(desc)
    thickness_mm = _extract_mingyang_thickness_mm(desc)
    thickness_mil = _extract_thickness_mil(desc, product)
    copper = _extract_copper(desc)
    foil = _extract_foil(desc) or "HTE"
    stack = _extract_stack(desc)
    length_in, width_in = _extract_size(desc, ignore_decimal=True)
    if not product or thickness_mm is None and thickness_mil is None or not copper or not stack or length_in is None or width_in is None:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "CCL规格缺少型号、厚度、铜厚、尺寸或配料结构")
    product_norm = _norm_product(product)
    product_rows = [row for row in rules.ccl_rows if row.product == product_norm]
    if not product_rows:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "明阳CCL胶系未找到")
    thickness_rows = [row for row in product_rows if _mingyang_thickness_matches(row, thickness_mil, thickness_mm)]
    if not thickness_rows:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "明阳CCL胶系找到，但厚度未匹配")
    stack_rows = [row for row in thickness_rows if row.stack == stack]
    if not stack_rows:
        stack_rows = [row for row in thickness_rows if _mingyang_stack_contains(row.stack, stack)]
    if not stack_rows:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"明阳CCL厚度找到，但配料结构未匹配：{stack}")

    exact_rows = [
        row
        for row in stack_rows
        if (row.copper == copper or row.copper == _reverse_copper(copper))
        and (not foil or not row.foil or row.foil == foil)
    ]
    for row in sorted(exact_rows, key=lambda item: item.excel_row):
        price_result = _mingyang_ccl_row_size_price(row, length_in, width_in)
        if price_result["ok"]:
            price = _round_money(price_result["price"])
            total = _calc_total(quantity, price)
            note = (
                f"命中明阳CCL报价 Sheet {row.sheet} 第 {row.excel_row} 行，"
                f"铜厚={row.copper}，铜箔={row.foil or '未写'}，尺寸列={price_result['label']}，公式={price_result['formula']}"
            )
            return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, price_result["label"])

    derived = _mingyang_derived_ccl_price(stack_rows, copper, foil, length_in, width_in)
    if derived["ok"]:
        price = _round_money(derived["price"])
        total = _calc_total(quantity, price)
        row = derived["row"]
        note = (
            f"命中明阳CCL说明2推导 Sheet {row.sheet} 第 {row.excel_row} 行，"
            f"目标铜厚={copper}，基准铜厚={row.copper}，基准41*49={derived['base_41']:.6g}，"
            f"尺寸={derived['size_label']}，公式={derived['formula']}"
        )
        return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, derived["size_label"])
    return ExtCalcResult("失败", "CCL", "待确认", "", "", "", derived.get("reason") or "明阳CCL报价行找到，但尺寸或说明2规则未覆盖")


def _extract_mingyang_thickness_mm(desc: str) -> float | None:
    match = re.search(r"(?<![A-Z])(\d+(?:\.\d+)?)\s*[±+/-]\s*\d+(?:\.\d+)?\s*MM\b", desc, re.I)
    if match:
        return float(match.group(1))
    return _extract_thickness_mm(desc)


def _mingyang_thickness_matches(row: ExtCclRule, thickness_mil: float | None, thickness_mm: float | None) -> bool:
    if _thickness_matches(row, thickness_mil, thickness_mm):
        return True
    if thickness_mm is not None and row.thickness_mm is not None and abs(row.thickness_mm - thickness_mm) <= 0.012:
        return True
    return False


def _mingyang_stack_contains(row_stack: str, target_stack: str) -> bool:
    row_pieces = _stack_piece_counts(row_stack)
    target_pieces = _stack_piece_counts(target_stack)
    return bool(target_pieces) and all(row_pieces.get(glass, 0) >= count for glass, count in target_pieces.items())


def _stack_piece_counts(stack: str) -> dict[str, int]:
    counts: dict[str, int] = {}
    for count, glass in re.findall(r"(\d+)\*([0-9]{3,4}|106)", stack):
        counts[glass] = counts.get(glass, 0) + int(count)
    return counts


def _extract_mingyang_pp_small_piece_mm(desc: str) -> tuple[float | None, float | None]:
    match = re.search(r"(\d+(?:\.\d+)?)\s*MM\s*[*xX×]\s*(\d+(?:\.\d+)?)\s*MM", desc, re.I)
    if not match:
        return None, None
    return float(match.group(1)) / 1000, float(match.group(2))


def _mingyang_ccl_row_size_price(row: ExtCclRule, length_in: float, width_in: float) -> dict[str, Any]:
    direct_key, direct_label = _direct_price_key(length_in, width_in, row.prices)
    if direct_key and row.prices.get(direct_key) is not None:
        price = float(row.prices[direct_key])
        return {"ok": True, "price": price, "label": direct_label, "formula": f"{price:.6g}"}
    parent = _select_parent("mingyang", length_in, width_in)
    if not parent:
        return {"ok": False, "reason": f"未找到可用尺寸父级：{length_in:g}*{width_in:g}"}
    parent_price = row.prices.get(parent["source_key"])
    if parent_price is None:
        return {"ok": False, "reason": f"尺寸父级{parent['source_label']}无报价"}
    price = float(parent_price) * parent["price_factor"] / parent["opens"]
    formula = f"{parent_price:.6g}*{parent['price_factor']:g}/{parent['opens']}"
    return {"ok": True, "price": price, "label": parent["label"], "formula": formula}


def _mingyang_derived_ccl_price(
    stack_rows: list[ExtCclRule],
    copper: str,
    foil: str,
    length_in: float,
    width_in: float,
) -> dict[str, Any]:
    base_copper, adjustment = _mingyang_ccl_base_adjustment(copper)
    if base_copper is None or adjustment is None:
        return {"ok": False, "reason": f"明阳CCL说明2未覆盖铜厚：{copper}"}
    if _mingyang_is_rtf_foil(foil) and _mingyang_rtf_needs_manual(copper):
        return {"ok": False, "reason": f"{copper}以上RTF铜价另议，未找到完全匹配报价行"}

    base_rows = [row for row in stack_rows if row.copper == base_copper]
    same_foil_rows = [row for row in base_rows if not foil or not row.foil or row.foil == foil]
    hte_rows = [row for row in base_rows if row.foil in {"", "HTE"}]
    if same_foil_rows:
        candidates = same_foil_rows
        foil_factor = 1.0
        foil_note = ""
    elif _mingyang_is_rtf_foil(foil) and hte_rows:
        candidates = hte_rows
        foil_factor = 1.03
        foil_note = "*1.03"
    else:
        return {"ok": False, "reason": f"未找到说明2基准铜厚行：{base_copper}/{foil or '未写铜箔'}"}

    size_factor = _mingyang_ccl_size_factor(length_in, width_in)
    if not size_factor:
        return {"ok": False, "reason": f"未找到可用尺寸换算：{length_in:g}*{width_in:g}"}
    for row in sorted(candidates, key=lambda item: item.excel_row):
        base_41 = row.prices.get("41")
        if base_41 is None:
            continue
        adjusted_41 = (float(base_41) + adjustment) * foil_factor
        price = adjusted_41 * size_factor["factor"]
        formula = f"({base_41:.6g}{adjustment:+.6g}){foil_note}*{size_factor['factor']:.6g}"
        return {
            "ok": True,
            "price": price,
            "row": row,
            "base_41": float(base_41),
            "size_label": size_factor["label"],
            "formula": formula,
        }
    return {"ok": False, "reason": f"说明2基准行缺少41*49价格：{base_copper}"}


def _mingyang_ccl_base_adjustment(copper: str) -> tuple[str | None, float | None]:
    normalized = copper.upper()
    if normalized in {"T/T", "H/H"}:
        return "H/H", 0.0
    if normalized == "J/J":
        return "H/H", -2.0
    if normalized == "2/2":
        return "H/H", 150.0
    if normalized in {"1/2", "2/1"}:
        return "H/H", 110.0
    if normalized == "1.5/1.5":
        return "H/H", 100.0
    if normalized in {"1/H", "H/1"}:
        return "1/1", -10.0
    if normalized == "3/3":
        return "H/H", 313.0
    return None, None


def _mingyang_ccl_size_factor(length_in: float, width_in: float) -> dict[str, Any] | None:
    dim = max(length_in, width_in)
    for target, factor in [(37, 0.9), (41, 1.0), (43, 1.05), (74, 1.8), (82, 2.0), (86, 2.1)]:
        if abs(dim - target) <= 0.8 and min(length_in, width_in) <= 49.8:
            return {"label": f"{target}*49", "factor": factor}
    parent = _select_parent("mingyang", length_in, width_in)
    if not parent:
        return None
    source_factor = {"37": 0.9, "41": 1.0, "43": 1.05}.get(parent["source_key"])
    if source_factor is None:
        return None
    factor = source_factor * parent["price_factor"] / parent["opens"]
    return {"label": parent["label"], "factor": factor}


def _mingyang_is_rtf_foil(foil: str) -> bool:
    return foil.upper().startswith("RTF")


def _mingyang_rtf_needs_manual(copper: str) -> bool:
    parts = [_to_float(part) for part in copper.replace("H", "0.5").split("/")]
    nums = [part for part in parts if part is not None]
    return bool(nums) and max(nums) >= 2


def _calculate_aoshikang_spec(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    parts = [_text(part) for part in desc.split("|")]
    first = parts[0] if parts else ""
    upper = desc.upper()
    if "PP" in first.upper() or (not first and "PP" in upper and "RC" in upper):
        return _calculate_aoshikang_pp(desc, rules)
    if "基板" in first or "CCL" in upper:
        return _calculate_aoshikang_ccl(desc, rules, quantity=quantity)
    return ExtCalcResult("失败", "未知", "待确认", "", "", "", "奥士康规格需包含基板或PP类型")


def _calculate_aoshikang_pp(desc: str, rules: ExtRules) -> ExtCalcResult:
    parts = [_text(part) for part in desc.split("|")]
    glass = _aoshikang_norm_glass(parts[1] if len(parts) > 1 else _extract_glass(desc))
    rc = _to_float(parts[2] if len(parts) > 2 else None)
    product = _aoshikang_extract_pp_product(parts, desc)
    if rc is not None and rc <= 1:
        rc *= 100
    if not product or not glass or rc is None:
        return ExtCalcResult("失败", "PP", "待确认", "", "", "", "奥士康PP关键字段缺失：胶系、玻布或RC")
    aliases = _aoshikang_pp_lookup_aliases(product)
    matches = [
        row
        for row in rules.pp_rows
        if row.product in aliases
        and row.glass == glass
        and row.rc_min is not None
        and row.rc_max is not None
        and row.rc_min - 0.001 <= rc <= row.rc_max + 0.001
    ]
    if not matches:
        product_rows = [row for row in rules.pp_rows if row.product in aliases]
        if not product_rows:
            reason = "PP胶系未找到"
        elif not [row for row in product_rows if row.glass == glass]:
            reason = "PP胶系找到，但玻布未找到"
        else:
            reason = "PP胶系和玻布找到，但RC未匹配"
        return ExtCalcResult("失败", "PP", "待确认", "", "", "", reason)
    best = sorted(matches, key=lambda row: (0 if row.product == product else 1, row.excel_row))[0]
    price = _aoshikang_precise_price(best.price)
    note = (
        f"命中奥士康PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，"
        f"胶系={best.product}，玻布={glass}，RC={rc:g}，未税每米价={price:.12g}"
    )
    return ExtCalcResult("成功", "PP", price, "", "", "", note, best.excel_row, best.sheet)


def _calculate_aoshikang_ccl(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    parsed = _parse_aoshikang_ccl_desc(desc)
    if parsed["error"]:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", parsed["error"])
    product = parsed["product"]
    product_options = _aoshikang_ccl_lookup_products(product)
    candidates = [row for row in rules.ccl_rows if row.product in product_options]
    if not candidates:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "CCL胶系未找到")
    candidates = [row for row in candidates if _aoshikang_thickness_matches(row, parsed)]
    if not candidates:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "CCL胶系找到，但厚度未找到")
    candidates = [row for row in candidates if _aoshikang_copper_matches(row.copper, parsed["copper"])]
    if not candidates:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"CCL厚度找到，但铜厚未匹配：{parsed['copper']}")
    preferred_kind = _aoshikang_preferred_kind(parsed["state"])
    if preferred_kind:
        kind_rows = [row for row in candidates if preferred_kind in row.kind or "总厚/芯厚" in row.kind]
        if kind_rows:
            candidates = kind_rows
    exact_stack_rows = [row for row in candidates if row.stack == parsed["stack"]]
    candidates = exact_stack_rows or [row for row in candidates if _aoshikang_stack_contains(row.stack, parsed["stack"])]
    if not candidates:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"CCL结构未匹配：{parsed['stack']}")
    exact_foil_rows = [
        row
        for row in candidates
        if row.foil == parsed["foil"] or (parsed["foil"].endswith("-CC") and row.foil == parsed["foil"].split("-", 1)[0])
    ]
    if exact_foil_rows:
        candidates = exact_foil_rows

    meta_map = (rules.ccl_notes or {}).get("products", {})
    last_reason = ""
    for row in sorted(candidates, key=lambda item: (0 if item.product == product else 1, item.excel_row)):
        meta = meta_map.get(row.product, _aoshikang_product_meta(row.product, bool(row.foil)))
        price_result = _aoshikang_ccl_row_price(row, parsed, meta)
        if price_result["ok"]:
            price = _aoshikang_precise_price(price_result["price"])
            total = _calc_total(quantity, price)
            note = (
                f"命中奥士康CCL报价 Sheet {row.sheet} 第 {row.excel_row} 行，"
                f"结构={row.stack}，铜厚={row.copper}，铜箔={parsed['foil'] or '未写'}，"
                f"尺寸列={price_result['size_label']}，公式={price_result['formula']}"
            )
            if price_result.get("foil_note"):
                note += f"；{price_result['foil_note']}"
            return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, price_result["size_label"])
        last_reason = price_result["reason"]
    return ExtCalcResult("失败", "CCL", "待确认", "", "", "", last_reason or "CCL报价行找到，但尺寸或铜箔未匹配")


def _calculate_taixing_spec(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    upper = desc.upper().strip()
    if upper.startswith("PREPREG"):
        return _calculate_taixing_pp(desc, rules)
    if upper.startswith("CCL"):
        return _calculate_taixing_ccl(desc, rules, quantity=quantity)
    return ExtCalcResult("失败", "未知", "待确认", "", "", "", "泰兴规格需以 Prepreg 或 CCL 开头")


def _calculate_taixing_pp(desc: str, rules: ExtRules) -> ExtCalcResult:
    product = _extract_product(desc)
    glass = _extract_glass(desc)
    rc = _extract_rc(desc)
    length = _extract_length(desc)
    if not product or not glass or rc is None:
        return ExtCalcResult("失败", "PP", "待确认", "", "", _fmt_length(length), "PP关键字段缺失：胶系、玻布或RC")
    product_norm = _norm_product(product)
    aliases = _taixing_pp_product_aliases(product_norm)
    matches: list[tuple[int, ExtPpRule]] = []
    for row in rules.pp_rows:
        if row.product not in aliases or row.glass != glass:
            continue
        if row.rc_min is not None and row.rc_max is not None and row.rc_min - 0.001 <= rc <= row.rc_max + 0.001:
            product_rank = 0 if row.product == product_norm else 1
            matches.append((product_rank, row))
    best = sorted(matches, key=lambda item: (item[0], item[1].excel_row))[0][1] if matches else None
    if not best:
        near_rows = [
            row
            for row in rules.pp_rows
            if row.product in aliases
            and row.glass == glass
            and row.rc_min is not None
            and row.rc_max is not None
            and abs(((row.rc_min + row.rc_max) / 2) - rc) <= 0.51
        ]
        best = sorted(near_rows, key=lambda row: (abs(((row.rc_min or 0) + (row.rc_max or 0)) / 2 - rc), row.excel_row))[0] if near_rows else None
    if not best:
        product_rows = [row for row in rules.pp_rows if row.product in aliases]
        if not product_rows:
            reason = "PP胶系未找到"
        elif not [row for row in product_rows if row.glass == glass]:
            reason = "PP胶系找到，但玻布未找到"
        else:
            reason = "PP胶系和玻布找到，但RC未匹配"
        return ExtCalcResult("失败", "PP", "待确认", "", "", _fmt_length(length), reason)
    if best.price is None:
        return ExtCalcResult("失败", "PP", "待确认", "", "", _fmt_length(length), "找到PP报价行，但每米基础价为空")
    price = _taixing_precise_price(float(best.price) * 1.04)
    note = (
        f"命中泰兴PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，"
        f"基础每米价={best.price:.6g}，公式={best.price:.6g}*1.04"
    )
    return ExtCalcResult("成功", "PP", price, "", _fmt_width(best.width), _fmt_length(length or best.length), note, best.excel_row, best.sheet)


def _calculate_taixing_ccl(desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    product = _extract_product(desc)
    thickness_mm = _extract_thickness_mm(desc)
    thickness_mil = _extract_thickness_mil(desc, product)
    copper = _extract_taixing_copper(desc)
    code = _extract_taixing_stack_code(desc)
    length_in, width_in = _extract_size(desc, ignore_decimal=False)
    if not product or thickness_mm is None or not copper or not code or length_in is None or width_in is None:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "CCL关键字段缺失：胶系、厚度、铜箔、尺寸或结构代码")
    product_norm = _norm_product(product)
    mapped_stack = _taixing_stack_from_code(code)
    mapped_stacks = {mapped_stack} if mapped_stack else set()
    if not mapped_stacks:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"CCL结构代码未配置：{code}")
    copper_group = _taixing_copper_group(product_norm, code, copper, length_in, width_in)
    if not copper_group:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"CCL铜箔价格组未匹配：{copper}")
    product_rows = [row for row in rules.ccl_rows if row.product == product_norm]
    if not product_rows:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "CCL胶系未找到")
    thickness_rows = [row for row in product_rows if _thickness_matches(row, thickness_mil, thickness_mm)]
    if not thickness_rows:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "CCL胶系找到，但厚度未找到")
    rows_to_try = [row for row in thickness_rows if row.stack in mapped_stacks]
    if not rows_to_try:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"CCL结构未匹配：{code}->{mapped_stack}")
    structure_note = f"；结构代码{code}->{mapped_stack}"
    family, side = _taixing_copper_family(copper)
    last_missing_reason = ""
    for row in rows_to_try:
        sheet_note = (rules.ccl_notes or {}).get(row.sheet, {})
        if family:
            special = _taixing_special_ccl_result(
                row,
                sheet_note,
                family,
                side,
                copper_group,
                length_in,
                width_in,
                quantity,
                structure_note,
            )
            if special.status == "成功":
                return special
            last_missing_reason = special.note
            continue
        regular = _taixing_regular_ccl_result(
            row,
            copper_group,
            length_in,
            width_in,
            quantity,
            structure_note,
        )
        if regular.status == "成功":
            return regular
        last_missing_reason = regular.note
    if last_missing_reason:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", last_missing_reason)
    return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "CCL胶系、厚度找到，但结构未匹配")


def _calculate_pp(customer_key: str, desc: str, rules: ExtRules) -> ExtCalcResult:
    product = _extract_product(desc)
    glass = _extract_glass(desc)
    rc = _extract_rc(desc)
    length = _extract_length(desc, floor_value=customer_key == "eaton")
    width = _extract_width(desc)
    small_length_m, small_width_in = _extract_pp_small_piece_size(desc)
    if not product or not glass or rc is None:
        return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width), _fmt_length(length), "PP规格缺少型号、玻布或RC")
    product_norm = _norm_product(product)
    base_product = product_norm[:-1] if product_norm.endswith("P") else product_norm
    products = {product_norm, base_product, f"{base_product}P"}
    if base_product.endswith("H"):
        products.add(f"{base_product[:-1]}P")
    matches: list[tuple[int, ExtPpRule]] = []
    for row in rules.pp_rows:
        if row.product not in products or row.glass != glass:
            continue
        if row.length is not None and length is not None and row.length != length:
            continue
        if row.rc_min is not None and row.rc_max is not None and row.rc_min - 0.001 <= rc <= row.rc_max + 0.001:
            product_rank = 0 if row.product == product_norm else 1
            matches.append((product_rank, row))
    if not matches and customer_key in {"eaton", "hanyu"}:
        for row in rules.pp_rows:
            if row.product not in products or row.glass != glass:
                continue
            if row.rc_min is not None and row.rc_max is not None and row.rc_min - 0.001 <= rc <= row.rc_max + 0.001:
                product_rank = 0 if row.product == product_norm else 1
                length_rank = abs((row.length or 0) - (length or row.length or 0))
                matches.append((product_rank + length_rank, row))
    best = sorted(matches, key=lambda item: (item[0], item[1].excel_row))[0][1] if matches else None
    if not best:
        return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width), _fmt_length(length), "未命中PP报价：型号、玻布、RC或卷长不匹配")
    if customer_key in {"eaton", "hanyu"} and small_length_m is not None and small_width_in is not None:
        if best.price is None:
            return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(small_width_in), "", "命中PP规格但缺少每米单价，无法计算小片")
        split = math.floor(49.5 / small_width_in) if small_width_in else 0
        if split <= 0:
            return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(small_width_in), "", f"PP小片纬向无法一开：{small_width_in:.2f} inch")
        raw_price = float(best.price) * small_length_m / split
        price = _hanyu_price(raw_price) if customer_key == "hanyu" else _round_money(raw_price)
        per_m_text = f"{best.price:.6f}" if customer_key == "hanyu" else f"{best.price:.2f}"
        price_text = f"{price:.6f}" if customer_key == "hanyu" else f"{price:.2f}"
        note = (
            f"PP小片命中 Sheet {best.sheet} 第 {best.excel_row} 行，Per M={per_m_text}，"
            f"经向={small_length_m:.3f}m，纬向={small_width_in:.2f}inch，纬向一开{split}，"
            f"公式={per_m_text}*{small_length_m:.3f}/{split}={price_text}"
        )
        return ExtCalcResult("成功", "PP", price, "", _fmt_width(small_width_in), "", note, best.excel_row, best.sheet)
    if customer_key == "hanyu" and length:
        if best.price is None:
            return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width), _fmt_length(length), "命中瀚宇PP规格但报价行缺少Per M")
        price = _hanyu_price(best.price)
        note = (
            f"命中瀚宇PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，"
            f"玻布={glass}，RC={rc:g}，卷长按{length}m匹配，Per M={price:.6f}"
        )
        return ExtCalcResult("成功", "PP", price, "", _fmt_width(width or best.width), _fmt_length(length), note, best.excel_row, best.sheet)
    if customer_key == "eaton" and length:
        if best.roll_price is not None:
            price = _round_money(best.roll_price)
            note = f"命中PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，每卷含税价={price:.2f}，卷长按{length}m匹配"
            return ExtCalcResult("成功", "PP", price, "", _fmt_width(width or best.width), _fmt_length(length), note, best.excel_row, best.sheet)
        if best.price is not None:
            price = _round_money(float(best.price) * length)
            note = f"命中PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，Per M={best.price:.2f}，公式={best.price:.2f}*{length}"
            return ExtCalcResult("成功", "PP", price, "", _fmt_width(width or best.width), _fmt_length(length), note, best.excel_row, best.sheet)
    if best.price is None:
        return ExtCalcResult("失败", "PP", "待确认", "", _fmt_width(width), _fmt_length(length), "命中PP规格但报价行缺少可用价格")
    price = _round_money(best.price)
    note = f"命中PP报价 Sheet {best.sheet} 第 {best.excel_row} 行，Per M={price:.2f}"
    if length:
        note += f"，卷长按{length}m匹配"
    return ExtCalcResult("成功", "PP", price, "", _fmt_width(width or best.width), _fmt_length(length), note, best.excel_row, best.sheet)


def _calculate_ccl(customer_key: str, desc: str, rules: ExtRules, quantity: Any = None) -> ExtCalcResult:
    product = _extract_product(desc)
    thickness_mil = _extract_thickness_mil(desc, product)
    thickness_mm = _extract_thickness_mm(desc)
    copper = _extract_copper(desc)
    foil = _extract_foil(desc) or ("HTE" if customer_key in {"hanyu", "eaton"} else "")
    stack = _extract_stack(desc)
    length_in, width_in = _extract_size(desc, ignore_decimal=customer_key in {"wutong", "eaton"})
    if not product or not copper or length_in is None or width_in is None:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "CCL规格缺少型号、铜厚或尺寸")
    if thickness_mil is None and thickness_mm is None:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "CCL规格缺少厚度")
    candidates = [
        row
        for row in rules.ccl_rows
        if row.product == _norm_product(product)
        and _thickness_matches(row, thickness_mil, thickness_mm)
        and (row.copper == copper or row.copper == _reverse_copper(copper))
        and (not stack or row.stack == stack)
        and (not row.foil or not foil or row.foil == foil)
    ]
    if not candidates and customer_key in {"eaton", "hanyu"} and stack:
        candidates = [
            row
            for row in rules.ccl_rows
            if row.product == _norm_product(product)
            and _thickness_matches(row, thickness_mil, thickness_mm)
            and (row.copper == copper or row.copper == _reverse_copper(copper))
            and (not row.foil or not foil or row.foil == foil)
        ]
    if not candidates:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "未命中基板报价：型号、厚度、铜厚、叠构或铜箔不匹配")

    if customer_key == "wutong":
        candidates = sorted(candidates, key=lambda row: (_thickness_distance(row, thickness_mil, thickness_mm), row.excel_row))

    if customer_key == "hanyu":
        return _calculate_hanyu_ccl_price(candidates, length_in, width_in, quantity)

    direct_key, direct_label = _direct_price_key(length_in, width_in, candidates[0].prices)
    for row in candidates:
        if direct_key and row.prices.get(direct_key) is not None:
            price = _round_money(row.prices[direct_key])
            total = _calc_total(quantity, price)
            return ExtCalcResult("成功", "CCL", price, total, "", "", f"命中基板报价 Sheet {row.sheet} 第 {row.excel_row} 行，尺寸列{direct_label}", row.excel_row, direct_label)

    parent = _select_parent(customer_key, length_in, width_in)
    if not parent:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"未找到可用小片父级：{length_in:g}*{width_in:g}")
    if customer_key == "eaton" and parent["opens"] > 1 and parent["price_factor"] > 1:
        for row in candidates:
            sf_price = row.prices.get("SF")
            if sf_price is None:
                continue
            area_length, area_width = _extract_size(desc, ignore_decimal=False)
            area_length = area_length or length_in
            area_width = area_width or width_in
            price = _round_money(float(sf_price) * area_length * area_width / 144)
            total = _calc_total(quantity, price)
            note = (
                f"基板按面积价命中 Sheet {row.sheet} 第 {row.excel_row} 行，SF={sf_price:.2f}，"
                f"尺寸={area_length:.2f}*{area_width:.2f}inch，公式={sf_price:.2f}*{area_length:.2f}*{area_width:.2f}/144"
            )
            return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, "SF")
    for row in candidates:
        parent_price = row.prices.get(parent["source_key"])
        if parent_price is None:
            continue
        price = _round_money(float(parent_price) * parent["price_factor"] / parent["opens"])
        total = _calc_total(quantity, price)
        note = (
            f"基板小片命中 Sheet {row.sheet} 第 {row.excel_row} 行，父级{parent['parent_w']}*{parent['parent_h']}，"
            f"按{parent['source_label']}价格{parent_price:.2f}*{parent['price_factor']}，"
            f"经向一开{parent['opens_w']}，纬向一开{parent['opens_h']}，总开数{parent['opens']}，"
            f"公式={parent_price:.2f}*{parent['price_factor']}/{parent['opens']}"
        )
        return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, parent["label"])
    if customer_key == "hanyu":
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "瀚宇CCL未找到尺寸列价格，不使用SF面积价")
    for row in candidates:
        sf_price = row.prices.get("SF")
        if sf_price is None:
            continue
        area_length, area_width = _extract_size(desc, ignore_decimal=False)
        area_length = area_length or length_in
        area_width = area_width or width_in
        price = _round_money(float(sf_price) * area_length * area_width / 144)
        total = _calc_total(quantity, price)
        note = (
            f"基板按面积价命中 Sheet {row.sheet} 第 {row.excel_row} 行，SF={sf_price:.2f}，"
            f"尺寸={area_length:.2f}*{area_width:.2f}inch，公式={sf_price:.2f}*{area_length:.2f}*{area_width:.2f}/144"
        )
        return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, "SF")
    return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"命中规格但父级尺寸列 {parent['source_label']} 无价格")


def _calculate_hanyu_ccl_price(
    candidates: list[ExtCclRule],
    length_in: float,
    width_in: float,
    quantity: Any = None,
) -> ExtCalcResult:
    forced = _hanyu_select_parent(length_in, width_in)
    if forced:
        for row in candidates:
            forced_price = row.prices.get(forced["source_key"])
            if forced_price is None:
                continue
            price = _round_money(forced_price)
            total = _calc_total(quantity, price)
            note = (
                f"命中瀚宇基板报价 Sheet {row.sheet} 第 {row.excel_row} 行，"
                f"客户尺寸={length_in:g}*{width_in:g}，强对应报价列={forced['source_key']}，"
                f"公式={forced_price:.2f}"
            )
            return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, forced["source_key"])
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "瀚宇CCL未找到强对应尺寸列价格，不使用SF面积价")

    direct_key, direct_label = _direct_price_key(length_in, width_in, candidates[0].prices)
    for row in candidates:
        if direct_key and row.prices.get(direct_key) is not None:
            price = _round_money(row.prices[direct_key])
            total = _calc_total(quantity, price)
            return ExtCalcResult("成功", "CCL", price, total, "", "", f"命中瀚宇基板报价 Sheet {row.sheet} 第 {row.excel_row} 行，尺寸列{direct_label}", row.excel_row, direct_label)

    parent = _select_parent("hanyu", length_in, width_in)
    if not parent:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"瀚宇CCL未找到尺寸列价格，不使用SF面积价：{length_in:g}*{width_in:g}")
    for row in candidates:
        parent_price = row.prices.get(parent["source_key"])
        if parent_price is None:
            continue
        price = _round_money(float(parent_price) * parent["price_factor"] / parent["opens"])
        total = _calc_total(quantity, price)
        note = (
            f"命中瀚宇基板报价 Sheet {row.sheet} 第 {row.excel_row} 行，"
            f"客户尺寸={length_in:g}*{width_in:g}，报价列={parent['source_key']}，"
            f"父级={parent['label']}，报价={parent_price:.2f}*{parent['price_factor']}，"
            f"经向一开{parent['opens_w']}，纬向一开{parent['opens_h']}，总开数{parent['opens']}，"
            f"公式={parent_price:.2f}*{parent['price_factor']}/{parent['opens']}"
        )
        return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, parent["source_key"])
    return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "瀚宇CCL未找到尺寸列价格，不使用SF面积价")


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\u3000", " ").strip()


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _text(value).replace(",", "").replace('"', "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else None


def _round_money(value: Any) -> float:
    return round(float(value) + 1e-9, 2)


def _taixing_precise_price(value: Any) -> float:
    return round(float(value), 6)


def _hanyu_price(value: Any) -> float:
    return round(float(value), 6)


def _taixing_ccl_price(value: Any) -> float:
    return round(float(value) + 0.00002, 2)


def _calc_total(quantity: Any, price: float | None) -> float | str:
    qty = _to_float(quantity)
    if qty is None or price is None:
        return ""
    return _round_money(qty * price)


def _norm_product(value: Any) -> str:
    text = _text(value).upper()
    text = re.sub(r"[\s_]+", "", text)
    return text.replace("-", "")


def _extract_product(desc: str) -> str:
    matches = re.findall(r"NY\s*-?\s*[A-Z]?\d{3,4}[A-Z0-9]*P?|NY\s*-?\s*A\d[A-Z0-9]*P?", desc, re.I)
    if not matches:
        return ""
    product = matches[-1]
    norm = _norm_product(product)
    return norm[:-1] if norm.endswith("P") and "RC" not in desc.upper() else norm


def _norm_glass(value: Any) -> str:
    text = _text(value).upper().replace("H", "")
    if re.search(r"(?<!\d)0+106(?!\d)", text):
        return "106"
    match = re.search(r"\d{3,4}|106", text)
    return match.group(0) if match else ""


def _norm_glasses(value: Any) -> list[str]:
    text = _text(value).upper().replace("H", "")
    normalized = []
    for item in re.findall(r"\d{3,4}|106", text):
        if re.fullmatch(r"0+106", item):
            normalized.append("106")
        else:
            normalized.append(item)
    return normalized


def _norm_taixing_glasses(value: Any) -> list[str]:
    text = _text(value).upper().replace("H", "")
    normalized = []
    for item in re.findall(r"\d{3,4}|106", text):
        if re.fullmatch(r"0+106", item):
            normalized.append("106")
        else:
            normalized.append(item)
    return normalized


def _extract_glass(desc: str) -> str:
    match = re.search(r"(?<!\d)(0106|1035|1067|1078|1080|1086|1037|1506|2113|2116|2313|3313|7628|106)H?(?!\d)", desc, re.I)
    return _norm_glass(match.group(1)) if match else ""


def _parse_rc_range(value: Any) -> tuple[float | None, float | None]:
    if isinstance(value, (int, float)):
        rc = float(value)
        rc = rc * 100 if rc <= 1 else rc
        return rc, rc
    text = _text(value)
    numbers = [float(item) for item in re.findall(r"\d+(?:\.\d+)?", text)]
    if not numbers:
        return None, None
    nums = [num * 100 if num <= 1 else num for num in numbers]
    if re.search(r"(?:<=|≤|≦)", text):
        return 0, max(nums)
    if re.search(r"(?:>=|≥|≧)", text):
        return min(nums), 999
    if ">" in text:
        return min(nums) + 0.001, 999
    if "<" in text:
        return 0, max(nums) - 0.001
    return min(nums), max(nums)


def _extract_rc(desc: str) -> float | None:
    match = re.search(r"RC\s*[:：]?\s*(\d+(?:\.\d+)?)", desc, re.I)
    if not match:
        return None
    value = float(match.group(1))
    return value * 100 if value <= 1 else value


def _length_int(value: Any) -> int | None:
    number = _to_float(value)
    return int(math.floor(number + 1e-9)) if number is not None else None


def _extract_length(desc: str, *, floor_value: bool = False) -> int | None:
    matches = re.findall(r"(?:[*xX×]\s*|\s)(\d+(?:\.\d+)?)\s*M\b", desc, re.I)
    if not matches:
        return None
    value = float(matches[-1])
    return int(math.floor(value + 1e-9)) if floor_value else int(round(value))


def _extract_width(desc: str) -> float | None:
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH|英寸)?\s*[*xX×]\s*(?:\d+(?:\.\d+)?)\s*M\b", desc, re.I)
    if match:
        return float(match.group(1))
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:\"|IN|INCH|英寸)\s*[*xX×]\s*\d+(?:\.\d+)?\s*M\b", desc, re.I)
    return float(match.group(1)) if match else None


def _norm_copper(value: Any) -> str:
    text = _text(value).upper().replace("OZ", "").replace(" ", "")
    text = text.replace("(RTF)", "").replace("(HTE)", "").replace("HO", "H").replace("H0", "H")
    if text == "T/T" or "(1/3)/(1/3)" in text:
        return "T/T"
    match = re.search(r"([0-9H.]+)\s*/\s*([0-9H.]+)", text)
    if not match:
        return ""
    if _looks_like_glass_pair(match.group(1), match.group(2)):
        return ""
    return f"{_norm_copper_part(match.group(1))}/{_norm_copper_part(match.group(2))}"


def _norm_copper_part(part: str) -> str:
    part = part.upper().replace("OZ", "")
    if part in {"H", "HO", "H0"}:
        return "H"
    number = _to_float(part)
    if number is None:
        return part
    return str(int(number)) if number.is_integer() else str(number).rstrip("0").rstrip(".")


def _reverse_copper(copper: str) -> str:
    parts = copper.split("/")
    return f"{parts[1]}/{parts[0]}" if len(parts) == 2 else copper


def _extract_copper(desc: str) -> str:
    if "(1/3)/(1/3)" in desc or "T/T" in desc.upper():
        return "T/T"
    match = re.search(
        r"(?<![A-Z0-9])([0-9Hh]+(?:\.\d+)?)\s*(?:\([^)]*\))?\s*/\s*([0-9Hh]+(?:\.\d+)?)\s*(?:\([^)]*\))?\s*(?:OZ)?",
        desc,
        re.I,
    )
    if not match:
        return ""
    return f"{_norm_copper_part(match.group(1))}/{_norm_copper_part(match.group(2))}"


def _norm_foil(value: Any) -> str:
    text = _text(value).upper()
    match = re.search(FOIL_TOKEN_PATTERN, text, re.I)
    return match.group(0).upper() if match else ""


def _extract_foil(desc: str) -> str:
    return _norm_foil(desc)


def _norm_stack(value: Any) -> str:
    text = _text(value).upper()
    pieces: list[tuple[str, int]] = []
    glass_pattern = r"1035|1067|1078|1080|1086|1506|2113|2116|2313|3313|7628|106"
    for count, glass in re.findall(rf"(?<!\d)(\d+)\s*(?:张|[*xX×])\s*({glass_pattern})(?!\d)", text):
        pieces.append((glass, int(count)))
    for glass, count in re.findall(rf"(?<!\d)({glass_pattern})\s*[*xX×]\s*(\d+)(?!\d)", text):
        pieces.append((glass, int(count)))
    if not pieces:
        return ""
    return "+".join(f"{count}*{glass}" for glass, count in sorted(pieces, key=lambda item: item[0]))


def _extract_stack(desc: str) -> str:
    return _norm_stack(desc)


def _extract_thickness_mil(desc: str, product: str = "") -> float | None:
    tol_match = re.search(r"(\d+(?:\.\d+)?)\s*[±+/-]\s*\d+(?:\.\d+)?\s*MIL", desc, re.I)
    if tol_match:
        return float(tol_match.group(1))
    match = re.search(r"(\d+(?:\.\d+)?(?:\s*\+\s*\d+(?:\.\d+)?)?)\s*MIL", desc, re.I)
    if match:
        return sum(float(part) for part in re.findall(r"\d+(?:\.\d+)?", match.group(1)))
    product_match = re.search(re.escape(product), _norm_product(desc), re.I) if product else None
    if product_match:
        tail = _norm_product(desc)[product_match.end() :]
        token = re.match(r"(\d{3}|[0-9]P[0-9])", tail)
        if token:
            return float(token.group(1).replace("P", ".").lstrip("0") or "0")
    match = re.search(r"\b(\d{3}|[0-9]P[0-9])\b", desc, re.I)
    if match:
        return float(match.group(1).upper().replace("P", ".").lstrip("0") or "0")
    return None


def _extract_thickness_mm(desc: str) -> float | None:
    match = re.search(r"(?<![A-Z])(\d+(?:\.\d+)?)\s*MM\b", desc, re.I)
    if match:
        return float(match.group(1))
    parts = [part for part in re.split(r"[_\s]+", desc) if part]
    for part in parts:
        if re.fullmatch(r"\d+(?:\.\d+)?", part):
            value = float(part)
            if 0 < value < 3:
                return value
    return None


def _extract_size(desc: str, *, ignore_decimal: bool = False) -> tuple[float | None, float | None]:
    match = re.search(
        r"(?:经向?|经)?\s*(\d+(?:\.\d+)?)\s*(MM|IN|INCH|英寸|\")?\s*[*xX×]\s*(?:纬向?|纬)?\s*(\d+(?:\.\d+)?)\s*(MM|IN|INCH|英寸|\")?",
        desc,
        re.I,
    )
    if not match:
        match = re.search(r"(?:经向?|经)\s*(\d+(?:\.\d+)?)\s*(MM|IN|INCH|英寸|\")?\s*[*xX×]?\s*(?:纬向?|纬)\s*(\d+(?:\.\d+)?)\s*(MM|IN|INCH|英寸|\")?", desc, re.I)
    if match:
        a, b = float(match.group(1)), float(match.group(3))
        unit_a = (match.group(2) or "").upper()
        unit_b = (match.group(4) or unit_a).upper()
        if not unit_a and not unit_b and int(b) in {106, 1067, 1078, 1080, 1506, 2113, 2116, 2313, 3313, 7628}:
            a = b = 9999
        is_mm = unit_a == "MM" or unit_b == "MM" or a > 120 or b > 120
        if is_mm:
            a, b = a / 25.4, b / 25.4
        if ignore_decimal:
            a, b = math.floor(a + 1e-9), math.floor(b + 1e-9)
        if a <= 120 and b <= 120:
            return a, b
    compact = re.search(r"\b(\d{4})\b", desc)
    compact_pairs = [item for item in re.findall(r"\b(\d{4})\b", desc) if int(item[:2]) <= 100 and int(item[2:]) <= 100]
    compact_pairs = [item for item in compact_pairs if item not in {"1067", "1078", "1080", "1506", "2113", "2116", "2313", "3313", "7628"}]
    if compact_pairs:
        compact = compact_pairs[-1]
        return float(compact[:2]), float(compact[2:])
    return None, None


def _thickness_matches(row: ExtCclRule, thickness_mil: float | None, thickness_mm: float | None) -> bool:
    if thickness_mil is not None and row.thickness_mil is not None and abs(row.thickness_mil - thickness_mil) <= 0.06:
        return True
    if thickness_mm is not None and row.thickness_mm is not None and abs(row.thickness_mm - thickness_mm) <= 0.006:
        return True
    if thickness_mil is not None and row.thickness_mm is not None and abs(row.thickness_mm - thickness_mil * 0.0254) <= 0.006:
        return True
    return False


def _thickness_distance(row: ExtCclRule, thickness_mil: float | None, thickness_mm: float | None) -> float:
    distances: list[float] = []
    if thickness_mil is not None and row.thickness_mil is not None:
        distances.append(abs(row.thickness_mil - thickness_mil) * 0.0254)
    if thickness_mm is not None and row.thickness_mm is not None:
        distances.append(abs(row.thickness_mm - thickness_mm))
    if thickness_mil is not None and row.thickness_mm is not None:
        distances.append(abs(row.thickness_mm - thickness_mil * 0.0254))
    return min(distances) if distances else 999.0


def _price_columns(headers: list[str]) -> dict[int, str]:
    price_cols: dict[int, str] = {}
    for idx, header in enumerate(headers, start=1):
        key = _price_key_from_label(header)
        if key:
            price_cols[idx] = key
    return price_cols


def _price_key_from_label(label: str) -> str:
    text = _text(label).upper().replace('"', "").replace("（含税）", "").replace("(含税)", "")
    if "SQ" in text or "SF" in text or text in {"SF", "SF单价"}:
        return "SF"
    numbers = re.findall(r"\d+(?:\.\d+)?", text)
    if len(numbers) >= 2 and abs(float(numbers[1]) - 49) <= 0.7:
        return _fmt_dim(float(numbers[0]))
    if len(numbers) >= 2:
        return f"{_fmt_dim(float(numbers[0]))}*{_fmt_dim(float(numbers[1]))}"
    return ""


def _row_prices(ws, row_idx: int, price_cols: dict[int, str]) -> dict[str, float | None]:
    prices: dict[str, float | None] = {}
    for col, key in price_cols.items():
        prices[key] = _to_float(ws.cell(row_idx, col).value)
    return prices


def _direct_price_key(length_in: float, width_in: float, prices: dict[str, float | None]) -> tuple[str, str]:
    full_a = f"{_fmt_dim(length_in)}*{_fmt_dim(width_in)}"
    full_b = f"{_fmt_dim(width_in)}*{_fmt_dim(length_in)}"
    if full_a in prices:
        return full_a, full_a
    if full_b in prices:
        return full_b, full_b
    if abs(width_in - 49) <= 0.8:
        key = _fmt_dim(length_in)
        if key in prices:
            return key, f"{key}*49"
    if abs(length_in - 49) <= 0.8:
        key = _fmt_dim(width_in)
        if key in prices:
            return key, f"{key}*49"
    return "", ""


def _select_parent(customer_key: str, length_in: float, width_in: float) -> dict | None:
    candidates = [
        (37, 49, "37", "37*49", 1),
        (41, 49, "41", "41*49", 1),
        (43, 49, "43", "43*49", 1),
        (74, 49, "37", "37*49*2", 2),
        (82, 49, "41", "41*49*2", 2),
        (86, 49, "43", "43*49*2", 2),
    ]
    valid = []
    for parent_w, parent_h, source_key, source_label, factor in candidates:
        opens_w = math.floor((parent_w + 1e-9) / length_in) if length_in else 0
        opens_h = math.floor((parent_h + 1e-9) / width_in) if width_in else 0
        opens = opens_w * opens_h
        if opens <= 0:
            continue
        fit_error = abs(length_in * opens_w - parent_w) + abs(width_in * opens_h - parent_h)
        valid.append(
            {
                "parent_w": parent_w,
                "parent_h": parent_h,
                "source_key": source_key,
                "source_label": source_label,
                "label": f"{parent_w}*{parent_h}",
                "price_factor": factor,
                "opens_w": opens_w,
                "opens_h": opens_h,
                "opens": opens,
                "fit_error": fit_error,
            }
        )
        if not valid:
            return None
    return sorted(valid, key=lambda item: (item["fit_error"], -item["parent_w"], -item["opens"]))[0]


def _hanyu_select_parent(length_in: float, width_in: float) -> dict | None:
    candidates = [
        (37, 49, "36*48", "37*49"),
        (41, 49, "40*48", "41*49"),
        (43, 49, "42*48", "43*49"),
    ]
    for expected_w, expected_h, source_key, source_label in candidates:
        if _same_size(length_in, width_in, expected_w, expected_h):
            return {
                "parent_w": expected_w,
                "parent_h": expected_h,
                "source_key": source_key,
                "source_label": source_label,
                "label": source_label,
                "price_factor": 1,
                "opens_w": 1,
                "opens_h": 1,
                "opens": 1,
            }
    return None


def _header_map(values: list[str]) -> dict[str, int]:
    counts: dict[str, int] = {}
    result: dict[str, int] = {}
    for idx, value in enumerate(values, start=1):
        if not value:
            continue
        counts[value] = counts.get(value, 0) + 1
        key = value if counts[value] == 1 else f"{value}__{counts[value]}"
        result[key] = idx
    return result


def _first_header_col(values: list[str], names: set[str]) -> int | None:
    for idx, value in enumerate(values, start=1):
        if value in names or any(name in value for name in names):
            return idx
    return None


def _find_header_contains(values: list[str], needles: set[str]) -> int | None:
    for idx, value in enumerate(values, start=1):
        normalized = value.replace("\n", "").replace(" ", "")
        if any(needle.upper() in normalized.upper() for needle in needles):
            return idx
    return None


def _looks_like_eaton_ccl_header(values: list[str]) -> bool:
    joined = "|".join(values)
    return bool(_find_header_contains(values, {"产品类别", "Products", "型号"})) and "SF" in joined and bool(
        _find_header_contains(values, {"结构", "组合结构"})
    )


def _looks_like_taixing_ccl_header(values: list[str]) -> bool:
    return bool(_find_header_contains(values, {"Product", "产品类别"})) and bool(_find_header_contains(values, {"结构", "组合"})) and any(
        _price_key_from_label(value) in {"36*48", "40*48", "42*48"} for value in values
    )


def _aoshikang_ccl_header_info(values: list[str]) -> dict[str, Any] | None:
    if not values or not values[0]:
        return None
    structure_col = _find_header_contains(values, {"结构"})
    price_col = _find_header_contains(values, {"新单价"})
    if not structure_col or not price_col:
        return None
    price_cols: dict[int, str] = {}
    for idx, header in enumerate(values, start=1):
        key = _aoshikang_price_key_from_label(header)
        if key:
            price_cols[idx] = key
    if not price_cols:
        return None
    kind_col = 4
    foil_col = None
    if len(values) >= 5 and "芯厚" in values[4]:
        kind_col = 5
        foil_col = 4
    elif len(values) >= 4 and "芯厚" in values[3]:
        kind_col = 4
        if len(values) >= 6 and not values[4] and "新单价" in values[5]:
            foil_col = 5
    return {"product": values[0], "kind_col": kind_col, "foil_col": foil_col, "price_cols": price_cols, "tax_col": price_col}


def _aoshikang_price_key_from_label(label: Any) -> str:
    text = _text(label).upper().replace('"', "")
    if "报价" not in text and "未税" not in text:
        return ""
    match = re.search(r"\d+(?:\.\d+)?", text)
    if not match:
        return ""
    number = int(math.floor(float(match.group(0)) + 1e-9))
    return str(number) if number in {37, 41, 43, 74, 80, 82, 86} else ""


def _aoshikang_norm_product(value: Any) -> str:
    text = _norm_product(value)
    aliases = {
        "TG150H": "NY2150H",
        "TG170": "NY2170",
        "2170H": "NY2170H",
        "无卤素TG150": "NY3150",
        "3170HC": "NY3170HC",
        "3170HF": "NY3170HF",
        "3170LK": "NY3170LK",
        "6200": "NY6200",
    }
    return aliases.get(text, text)


def _aoshikang_product_aliases(value: Any) -> list[str]:
    raw = _text(value)
    norm = _aoshikang_norm_product(raw)
    aliases: set[str] = set()
    explicit = {
        "NY3150": {"NY3150", "NY3150HF", "NY3150HC"},
        "无卤素TG150": {"NY3150", "NY3150HF", "NY3150HC"},
        "TG170": {"NY2170"},
        "TG150H": {"NY2150H"},
        "2170H": {"NY2170H"},
        "3170HC": {"NY3170HC"},
        "3170HF": {"NY3170HF"},
        "3170LK": {"NY3170LK"},
        "NYP1": {"NYP1"},
        "NYP2": {"NYP2"},
        "NYP3": {"NYP3"},
        "NYP4": {"NYP4"},
    }
    aliases.update(explicit.get(norm, set()))
    aliases.add(norm)
    for token in re.split(r"[/、&]+", raw):
        token_norm = _aoshikang_norm_product(token)
        if token_norm:
            aliases.add(token_norm)
    return sorted(alias for alias in aliases if alias)


def _aoshikang_pp_segment_aliases(label: Any) -> set[str]:
    aliases: set[str] = set()
    raw_label = _text(label)
    text = raw_label.upper().replace(" ", "")
    if not any(marker in text for marker in ("NY", "TG", "2170", "3170", "6200", "6666", "9999")):
        return aliases
    for token in re.split(r"[/、&]+", raw_label):
        token_norm = _aoshikang_norm_product(token)
        if not token_norm:
            continue
        aliases.add(token_norm if token_norm.endswith("P") else f"{token_norm}P")
    if "NY3150HF" in text or "NY3150H F" in text or "NY3150HF/NY3150HC" in text or "NY3150" in text and "HF" in text:
        aliases.update({"NY3150HFP", "NY3150HCP"})
    if "NY3170M" in text:
        aliases.add("NY3170MP")
    if "NY3170HC" in text:
        aliases.add("NY3170HCP")
    if "NY3170HF" in text:
        aliases.add("NY3170HFP")
    return aliases


def _aoshikang_pp_special_products(header: Any) -> set[str]:
    text = _text(header).upper().replace(" ", "").replace("-", "")
    if "NYA1" in text:
        return {"NYA1P"}
    if "NYA2" in text:
        return {"NYA2P"}
    return set()


def _aoshikang_extract_pp_product(parts: list[str], desc: str) -> str:
    candidates: list[str] = []
    if len(parts) > 4:
        candidates.append(parts[4])
    candidates.extend(parts[3:])
    candidates.append(desc)
    for value in candidates:
        normalized = _aoshikang_norm_product(value)
        if re.fullmatch(r"NY[A-Z0-9]+P?", normalized):
            return normalized
        product = _extract_product(_text(value))
        if product:
            return _aoshikang_norm_product(product)
    return ""


def _aoshikang_pp_lookup_aliases(product: str) -> set[str]:
    product = _aoshikang_norm_product(product)
    aliases = {product}
    if product.endswith("P"):
        aliases.add(product[:-1])
    else:
        aliases.add(f"{product}P")
    explicit = {
        "NY6300SNP": {"NY6300SNP", "NY6300SN"},
        "NY6300SN": {"NY6300SNP", "NY6300SN"},
        "NY6180LP": {"NY6180LP", "NY6180L"},
        "NY6180L": {"NY6180LP", "NY6180L"},
        "NY3170MP": {"NY3170MP", "NY3170M"},
        "NY3170M": {"NY3170MP", "NY3170M"},
        "NY3170M2P": {"NY3170M2P", "NY3170M2"},
        "NY3170M2": {"NY3170M2P", "NY3170M2"},
        "NY3176HFP": {"NY3176HFP", "NY3176HF"},
        "NY3176HF": {"NY3176HFP", "NY3176HF"},
    }
    aliases.update(explicit.get(product, set()))
    return aliases


def _aoshikang_ccl_lookup_products(product: str) -> list[str]:
    product = _aoshikang_norm_product(product)
    if product == "NYA1":
        return ["NYA1", "NY2150"]
    if product == "NYA2":
        return ["NYA2", "NY2170"]
    return [product]


def _aoshikang_product_meta(product: str, has_foil_col: bool) -> dict[str, Any]:
    product = _aoshikang_norm_product(product)
    if has_foil_col:
        return {"strict_foil": True}
    if product == "NY3170M":
        return {"default_foil": "RTF", "ny3170m": True}
    if product in {"NY2140", "NY2140L", "NY2150", "NY2150H", "NY2170", "NY2170H", "NY3150", "NY3150HF", "NY3150HC", "NY1600", "NYA1", "NYA2"}:
        return {"default_foil": "HTE", "rtf_from_hte": True}
    return {"strict_foil": False}


def _aoshikang_norm_glass(value: Any) -> str:
    text = _text(value).upper().replace(" ", "")
    if re.fullmatch(r"0+106", text):
        return "106"
    match = re.search(r"7628T|1027|1035|1037|1067|1078|1080|1086|1506|2113|2116|2313|3313|7628|106", text)
    glass = match.group(0) if match else ""
    return "7628" if glass == "7628T" else glass


def _aoshikang_norm_copper(value: Any) -> str:
    text = _text(value).upper().replace(" ", "").replace("\\", "/")
    aliases = {
        "J/J": "J/J",
        "R/R": "1.5/1.5",
        "X/1": "0/1",
        "1/X": "1/0",
    }
    if text in aliases:
        return aliases[text]
    return _norm_copper(text)


def _aoshikang_copper_matches(row_copper: str, spec_copper: str) -> bool:
    if row_copper == spec_copper:
        return True
    return {row_copper, spec_copper} == {"0/1", "1/0"}


def _aoshikang_thickness_matches(row: ExtCclRule, parsed: dict[str, Any]) -> bool:
    thickness = parsed.get("thickness")
    if thickness is None or row.thickness_mm is None:
        return False
    tolerance = 0.0005 if parsed.get("product") == "NY3170M2" else 0.006
    return abs(row.thickness_mm - float(thickness)) <= tolerance


def _aoshikang_norm_foil(value: Any) -> str:
    text = _text(value).upper()
    base = _norm_foil(text)
    if base and "长春" in text:
        return f"{base}-CC"
    return base


def _aoshikang_norm_stack(value: Any) -> str:
    text = _text(value).upper().replace("×", "*").replace("X", "*").replace("/", "+")
    text = re.sub(r"\s+", "", text)
    glass_pattern = r"7628T|1027|1035|1037|1067|1078|1080|1086|1506|2113|2116|2313|3313|7628|106"
    pieces: list[tuple[str, int]] = []
    for count, glass in re.findall(rf"(?<!\d)(\d+)\s*(?:张|\*)\s*({glass_pattern})(?!\d)", text):
        pieces.append((glass, int(count)))
    for glass, count in re.findall(rf"(?<!\d)({glass_pattern})\s*\*\s*(\d+)(?!\d)", text):
        pieces.append((glass, int(count)))
    if not pieces:
        single = re.fullmatch(rf"({glass_pattern})", text)
        if single:
            pieces.append((single.group(1), 1))
    if not pieces:
        return ""
    totals: dict[str, int] = {}
    for glass, count in pieces:
        totals[glass] = totals.get(glass, 0) + count
    return "+".join(f"{totals[glass]}*{glass}" for glass in sorted(totals))


def _aoshikang_stack_contains(row_stack: str, spec_stack: str) -> bool:
    row_counts = _stack_counts(row_stack)
    spec_counts = _stack_counts(spec_stack)
    if not row_counts or not spec_counts:
        return False
    return all(row_counts.get(glass, 0) >= count for glass, count in spec_counts.items())


def _stack_counts(stack: str) -> dict[str, int]:
    counts: dict[str, int] = {}
    for count, glass in re.findall(r"(\d+)\*([A-Z0-9]+)", stack):
        counts[glass] = counts.get(glass, 0) + int(count)
    return counts


def _aoshikang_preferred_kind(state: str) -> str:
    text = _text(state)
    if "不含铜" in text:
        return "芯厚"
    if "含铜" in text:
        return "总厚"
    return ""


def _parse_aoshikang_ccl_desc(desc: str) -> dict[str, Any]:
    desc = _aoshikang_strip_tolerance(desc)
    parts = [_text(part) for part in desc.split("|")]
    product = _aoshikang_norm_product(parts[1] if len(parts) > 1 else _extract_product(desc))
    thickness = _to_float(parts[2] if len(parts) > 2 else None)
    copper = _aoshikang_norm_copper(parts[3] if len(parts) > 3 else desc)
    length_in = width_in = None
    if len(parts) > 4:
        length_in, width_in = _aoshikang_parse_size(parts[4])
    if length_in is None or width_in is None:
        length_in, width_in = _extract_size(desc, ignore_decimal=False)
    foil = _aoshikang_norm_foil(parts[7] if len(parts) > 7 else _extract_foil(desc)) or "HTE"
    stack = _aoshikang_norm_stack(parts[11] if len(parts) > 11 else desc)
    state = parts[5] if len(parts) > 5 else ""
    error = ""
    if not product or thickness is None or not copper or length_in is None or width_in is None or not stack:
        error = "奥士康CCL关键字段缺失：胶系、厚度、铜厚、尺寸或结构"
    return {
        "product": product,
        "thickness": thickness,
        "copper": copper,
        "length": length_in,
        "width": width_in,
        "foil": foil,
        "stack": stack,
        "state": state,
        "error": error,
    }


def _aoshikang_strip_tolerance(desc: str) -> str:
    text = _text(desc)
    return re.sub(r"\s*[+＋]\s*\d+(?:\.\d+)?\s*/\s*[-－]\s*\d+(?:\.\d+)?\s*$", "", text)


def _aoshikang_parse_size(value: Any) -> tuple[float | None, float | None]:
    text = _text(value).upper().replace("Ｘ", "X").replace("×", "X").replace("*", "X")
    match = re.search(r"(\d+(?:\.\d+)?)\s*X\s*(\d+(?:\.\d+)?)", text)
    if not match:
        return None, None
    return float(match.group(1)), float(match.group(2))


def _aoshikang_size_price(row: ExtCclRule, length_in: float, width_in: float) -> tuple[float | None, str, str]:
    length_floor = int(math.floor(length_in + 1e-9))
    width_floor = int(math.floor(width_in + 1e-9))
    if {length_floor, width_floor} == {41, 43}:
        base = row.prices.get("41")
        if base is None:
            return None, "41", "41*43特殊尺寸缺少41列"
        return float(base) * 43 / 49 * 1.08, "41*43", f"{base:.12g}*43/49*1.08"
    if {length_floor, width_floor} == {37, 43}:
        base = row.prices.get("41")
        if base is None:
            return None, "41", "37*43特殊尺寸缺少41列"
        return float(base) * 43 / 48 * 0.9 * 1.08, "37*43", f"{base:.12g}*43/48*0.9*1.08"
    candidates = [37, 41, 43, 74, 80, 82, 86]
    if abs(width_in - 49) <= 1.0 or abs(width_in - 46) <= 1.0:
        major = length_in
    elif abs(length_in - 49) <= 1.0 or abs(length_in - 46) <= 1.0:
        major = width_in
    else:
        major = max(length_in, width_in)
    key = str(min(candidates, key=lambda value: abs(value - major)))
    if key not in row.prices:
        key = str(length_floor if str(length_floor) in row.prices else width_floor)
    price = row.prices.get(key)
    if price is None:
        return None, key, f"找到CCL报价行，但{key}列无价格"
    return float(price), key, f"{price:.12g}"


def _aoshikang_ccl_row_price(row: ExtCclRule, parsed: dict[str, Any], meta: dict[str, Any]) -> dict[str, Any]:
    spec_foil = parsed["foil"]
    if row.foil:
        if row.foil != spec_foil and not (spec_foil.endswith("-CC") and row.foil == spec_foil.split("-", 1)[0]):
            return {"ok": False, "reason": f"报价行铜箔={row.foil}，未匹配规格铜箔={spec_foil}"}
        foil_note = "行内铜箔精确匹配" if row.foil == spec_foil else "未找到同厚度长春铜箔，按同类HTE行匹配"
        adjusted = None
    else:
        default_foil = meta.get("default_foil", "")
        if meta.get("strict_foil") and spec_foil:
            return {"ok": False, "reason": f"该报价Sheet需行内铜箔精确匹配，未找到{spec_foil}"}
        adjusted = _aoshikang_default_foil_adjustment(meta, spec_foil)
        if adjusted is None:
            return {"ok": False, "reason": f"报价单说明未覆盖铜箔：{spec_foil}"}
        foil_note = adjusted["note"]
        if not default_foil:
            foil_note = "报价行未限定铜箔"
    base_price, size_label, formula = _aoshikang_size_price(row, parsed["length"], parsed["width"])
    if base_price is None:
        return {"ok": False, "reason": formula}
    final_price = base_price
    if adjusted:
        kind = adjusted["kind"]
        if kind == "minus" and row.product == "NY3170M" and meta.get("ny3170m") and spec_foil == "HTE":
            special = _aoshikang_ny3170m_hte_tax_price(row, parsed)
            if special:
                return special
        if kind == "plus" and row.product == "NY3170M" and meta.get("ny3170m") and spec_foil == "RTF2":
            special = _aoshikang_ny3170m_rtf2_base41_price(row, parsed)
            if special:
                return special
        if kind == "rtf_from_hte":
            final_price = base_price * 1.05
            formula = f"{formula}*1.05"
        elif kind == "minus":
            final_price = base_price + adjusted["amount"]
            formula = f"{formula}{adjusted['amount']:+g}"
        elif kind == "plus":
            final_price = base_price + adjusted["amount"]
            formula = f"{formula}+{adjusted['amount']:g}"
        elif kind == "multiplier":
            final_price = base_price * adjusted["multiplier"]
            formula = f"{formula}*{adjusted['multiplier']:g}"
    if parsed["product"] == "NYA1" and row.product == "NY2150":
        final_price *= 1.05
        formula = f"({formula})*1.05"
        foil_note = f"{foil_note}；NY-A1按NY2150基础价*1.05"
    if parsed["product"] == "NYA2" and row.product == "NY2170":
        final_price *= 1.05
        formula = f"({formula})*1.05"
        foil_note = f"{foil_note}；NY-A2按NY2170基础价*1.05"
    return {"ok": True, "price": final_price, "size_label": size_label, "formula": formula, "foil_note": foil_note}


def _aoshikang_ny3170m_hte_tax_price(row: ExtCclRule, parsed: dict[str, Any]) -> dict[str, Any] | None:
    tax_price = row.prices.get("TAX")
    if tax_price is None:
        return None
    key = _aoshikang_standard_size_key(parsed["length"], parsed["width"])
    if not key:
        return None
    ratio = round(float(key) / 41, 2)
    price = (float(tax_price) - 2) / 1.13 * ratio
    formula = f"({tax_price:.12g}-2)/1.13*{ratio:.2f}"
    note = f"NY3170M默认RTF，HTE=RTF-2；{key}/41保留两位={ratio:.2f}"
    return {"ok": True, "price": price, "size_label": f"{key}/41", "formula": formula, "foil_note": note}


def _aoshikang_ny3170m_rtf2_base41_price(row: ExtCclRule, parsed: dict[str, Any]) -> dict[str, Any] | None:
    base_41 = row.prices.get("41")
    if base_41 is None:
        return None
    key = _aoshikang_standard_size_key(parsed["length"], parsed["width"])
    if not key:
        return None
    ratio = round(float(key) / 41, 2)
    price = (float(base_41) + 15) * ratio
    formula = f"({base_41:.12g}+15)*{ratio:.2f}"
    note = f"NY3170M行未标明RTF2，按RTF2=RTF+15；41*49基础价={base_41:.12g}；{key}/41保留两位={ratio:.2f}"
    return {"ok": True, "price": price, "size_label": f"41->{key}/41", "formula": formula, "foil_note": note}


def _aoshikang_standard_size_key(length_in: float, width_in: float) -> str:
    candidates = [37, 41, 43, 74, 80, 82, 86]
    if abs(width_in - 49) <= 1.0 or abs(width_in - 46) <= 1.0:
        major = length_in
    elif abs(length_in - 49) <= 1.0 or abs(length_in - 46) <= 1.0:
        major = width_in
    else:
        major = max(length_in, width_in)
    return str(min(candidates, key=lambda value: abs(value - major)))


def _aoshikang_default_foil_adjustment(meta: dict[str, Any], spec_foil: str) -> dict[str, Any] | None:
    default_foil = meta.get("default_foil", "")
    if not default_foil:
        return {"kind": "base", "note": "报价行未限定铜箔"}
    if spec_foil == default_foil:
        return {"kind": "base", "note": f"默认{default_foil}铜箔"}
    if meta.get("rtf_from_hte") and default_foil == "HTE" and spec_foil == "RTF":
        return {"kind": "rtf_from_hte", "note": "RTF按HTE*1.05计算，不取整"}
    if meta.get("ny3170m"):
        if spec_foil == "HTE":
            return {"kind": "minus", "amount": -2.0, "note": "NY3170M按HTE=RTF-2"}
        if spec_foil == "RTF2":
            return {"kind": "plus", "amount": 15.0, "note": "NY3170M按RTF2=RTF+15"}
    return None


def _aoshikang_precise_price(value: Any) -> float:
    return round(float(value), 12)


def _taixing_ccl_price_columns(ws, header_row: int, headers: list[str]) -> dict[int, str]:
    price_cols: dict[int, str] = {}
    current_group = ""
    for col_idx, header in enumerate(headers, start=1):
        group_label = _text(ws.cell(header_row - 1, col_idx).value)
        if group_label:
            current_group = _taixing_price_group_from_label(group_label) or current_group
        size_key = _price_key_from_label(header)
        if current_group and size_key in {"SF", "36*48", "40*48", "42*48"}:
            price_cols[col_idx] = f"{current_group}:{size_key}"
    return price_cols


def _taixing_price_group_from_label(label: Any) -> str:
    text = _text(label).upper().replace(" ", "")
    if "2OZ" in text or "2/2" in text or "H/2" in text:
        return "2OZ"
    if "1OZ" in text or "1/1" in text or "H/1" in text or "HOZ/1OZ" in text:
        return "1OZ"
    if "HOZ/HOZ" in text or "HOZ" in text or "Hoz/Hoz".upper() in text:
        return "HOZ"
    return ""


def _taixing_pp_product_aliases(product: str) -> set[str]:
    product = _norm_product(product)
    aliases = {product}
    if product.endswith("P"):
        aliases.add(product[:-1])
    else:
        aliases.add(f"{product}P")
    return aliases


def _extract_taixing_copper(desc: str) -> str:
    match = re.search(r"\b(R2H|SH|S1|S2|R21|R31|HOZ|H)\s*/\s*(R2H|SH|S1|S2|R21|R31|HOZ|H)\b", desc, re.I)
    if not match:
        return ""
    left = match.group(1).upper()
    right = match.group(2).upper()
    if left == "HOZ":
        left = "H"
    if right == "HOZ":
        right = "H"
    return f"{left}/{right}"


def _extract_taixing_stack_code(desc: str) -> str:
    match = re.search(r"\b(\d+X\d{3,4}|\d+[A-Z])\b\s+NY", desc, re.I)
    if not match:
        match = re.search(r"\b(\d+X\d{3,4}|\d+[A-Z])\b", desc, re.I)
    return match.group(1).upper() if match else ""


def _taixing_stack_from_code(code: str) -> str:
    text = _text(code).upper().replace("×", "X").replace("*", "X")
    explicit = re.fullmatch(r"(\d+)X(\d{3,4})", text)
    if explicit:
        return _norm_taixing_stack(f"{explicit.group(1)}*{explicit.group(2)}")
    match = re.fullmatch(r"(\d+)([A-Z])", text)
    if not match:
        return ""
    glass_map = {
        "A": "106",
        "B": "1065",
        "C": "1067",
        "D": "1078",
        "E": "1080",
        "F": "1086",
        "G": "2112",
        "H": "2113",
        "I": "2313",
        "J": "3313",
        "K": "2116",
        "L": "2165",
        "M": "1500",
        "N": "1501",
        "O": "1504",
        "P": "1506",
        "Q": "1652",
        "R": "6700",
        "S": "7627",
        "T": "7628",
        "U": "7629",
        "V": "7530",
        "X": "1037",
    }
    glass = glass_map.get(match.group(2))
    return _norm_taixing_stack(f"{match.group(1)}*{glass}") if glass else ""


def _norm_taixing_stack(value: Any) -> str:
    text = _text(value).upper().replace("×", "*").replace("X", "*").replace("＊", "*")
    text = re.sub(r"\s+", "", text)
    glass_pattern = r"1035|1037|1065|1067|1078|1080|1086|1500|1501|1504|1506|1652|2112|2113|2116|2165|2313|3313|6700|7530|7627|7628|7629|106"
    pieces: list[tuple[str, int]] = []
    for count, glass in re.findall(rf"(?<!\d)(\d+)\*({glass_pattern})(?!\d)", text):
        pieces.append((glass, int(count)))
    for glass, count in re.findall(rf"(?<!\d)({glass_pattern})\*(\d+)(?!\d)", text):
        pieces.append((glass, int(count)))
    if not pieces:
        return ""
    totals: dict[str, int] = {}
    order: list[str] = []
    for glass, count in pieces:
        if glass not in totals:
            order.append(glass)
            totals[glass] = 0
        totals[glass] += count
    return "+".join(f"{totals[glass]}*{glass}" for glass in order)


def _taixing_copper_group(product: str, code: str, copper: str, length_in: float, width_in: float) -> str:
    normalized = copper.upper()
    reversed_copper = _reverse_copper(normalized)
    if product == "NY6300S" and code.upper() == "2X1035" and normalized == "S2/S2" and _same_size(length_in, width_in, 41, 49):
        return "1OZ"
    if normalized in {"SH/SH", "H/H", "R2H/R2H"}:
        return "HOZ"
    if normalized in {"SH/S1", "S1/SH", "S1/S1", "R21/R21", "R31/R31", "R2H/R21", "R21/R2H"} or reversed_copper in {"SH/S1", "S1/SH", "R2H/R21", "R21/R2H"}:
        return "1OZ"
    if normalized in {"S2/S2", "S1/S2", "S2/S1", "SH/S2", "S2/SH", "H/S2", "S2/H"} or reversed_copper in {
        "S1/S2",
        "S2/S1",
        "SH/S2",
        "S2/SH",
        "H/S2",
        "S2/H",
    }:
        return "2OZ"
    return ""


def _parse_taixing_ccl_sheet_notes(ws) -> dict[str, Any]:
    lines: list[str] = []
    for row_idx in range(1, ws.max_row + 1):
        values = [_text(ws.cell(row_idx, col).value) for col in range(1, ws.max_column + 1)]
        line = " ".join(value for value in values if value)
        if line and re.search(r"RTF|HVLP|HTE|另算|另议|价格", line, re.I):
            lines.append(line)
    note_text = "；".join(lines)
    parsed: dict[str, Any] = {
        "text": note_text,
        "amounts": {},
        "percents": {},
        "base_families": set(),
        "manual_families": set(),
        "same_families": set(),
    }
    for line in lines:
        upper = line.upper().replace(" ", "")
        main_family = _taixing_note_main_family(upper)
        if not main_family:
            continue
        if "以上价格" in line or "价格为" in line:
            parsed["base_families"].add(main_family)
        if ("另算" in line or "另议" in line) and "其它" not in line and "其他" not in line:
            parsed["manual_families"].add(main_family)
        if "同价" in line:
            parsed["same_families"].add(main_family)
        percent = _taixing_note_percent(upper)
        if percent is not None:
            parsed["percents"][main_family] = percent
        for group, side, amount in _taixing_note_sf_amounts(line):
            parsed["amounts"][(main_family, group, side)] = amount
    parsed["has_special"] = bool(
        parsed["amounts"] or parsed["percents"] or parsed["base_families"] or parsed["manual_families"] or parsed["same_families"]
    )
    return parsed


def _taixing_note_main_family(text: str) -> str:
    for pattern in (
        r"使用(RTF4|RTF3|RTF2|RTF|HVLP[1-4]?)",
        r"(RTF4|RTF3|RTF2|RTF|HVLP[1-4]?)铜箔",
        r"以上价格为(RTF4|RTF3|RTF2|RTF|HVLP[1-4]?)",
        r"(RTF4|RTF3|RTF2|RTF|HVLP[1-4]?)",
    ):
        match = re.search(pattern, text, re.I)
        if match:
            return match.group(1).upper()
    return ""


def _taixing_note_percent(text: str) -> float | None:
    match = re.search(r"(上调|增加|加|下调|减|减少)\s*(\d+(?:\.\d+)?)%", text)
    if not match:
        return None
    value = float(match.group(2)) / 100
    return -value if match.group(1) in {"下调", "减", "减少"} else value


def _taixing_note_sf_amounts(line: str) -> list[tuple[str, str, float]]:
    normalized = line.upper().replace("，", ",").replace("；", ";")
    groups = ["HOZ", "1OZ", "2OZ"]
    found: list[tuple[str, str, float]] = []
    for group in groups:
        start = normalized.find(group)
        if start < 0:
            continue
        next_starts = [normalized.find(other, start + len(group)) for other in groups if other != group]
        next_starts = [idx for idx in next_starts if idx >= 0]
        end = min(next_starts) if next_starts else len(normalized)
        segment = normalized[start:end]
        for side_name, side_key in (("单面", "single"), ("双面", "double")):
            match = re.search(rf"{side_name}\s*(增加|加|减|减少)?\s*(\d+(?:\.\d+)?)", segment)
            if not match:
                continue
            amount = float(match.group(2))
            if match.group(1) in {"减", "减少"}:
                amount = -amount
            found.append((group, side_key, amount))
    return found


def _taixing_copper_family(copper: str) -> tuple[str, str]:
    normalized = copper.upper()
    if normalized in {"R21/R21", "R2H/R2H"}:
        return "RTF2", "double"
    if normalized == "R31/R31":
        return "RTF3", "double"
    if normalized in {"R2H/R21", "R21/R2H"}:
        return "RTF2", "mixed"
    return "", ""


def _taixing_special_ccl_result(
    row: ExtCclRule,
    sheet_note: dict[str, Any],
    family: str,
    side: str,
    copper_group: str,
    length_in: float,
    width_in: float,
    quantity: Any,
    structure_note: str,
) -> ExtCalcResult:
    if not sheet_note or not sheet_note.get("has_special"):
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "需人工确认：报价单说明未覆盖该情况")
    if family in sheet_note.get("manual_families", set()):
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "需人工确认：报价单说明未覆盖该情况")

    adjustment_type, adjustment = _taixing_sheet_adjustment(sheet_note, family, copper_group, side)
    if adjustment_type in {"same_price", "base_price"}:
        result = _taixing_regular_ccl_result(row, copper_group, length_in, width_in, quantity, structure_note)
        if result.status == "成功":
            label = "同价" if adjustment_type == "same_price" else "基础价"
            result.note = result.note.replace("公式=", f"报价说明={family}{label}，按尺寸列取价，公式=")
        return result
    if adjustment_type == "explicit_sf_amount" and adjustment is not None:
        sf_key = f"{copper_group}:SF"
        base_sf = row.prices.get(sf_key)
        if base_sf is None:
            return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"找到CCL报价行，但{copper_group}组SF列无价格")
        parent = _taixing_ccl_sf_parent(length_in, width_in)
        if not parent:
            return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "需人工确认：报价单说明未覆盖该情况")
        multiplier = _taixing_ccl_multiplier(length_in, width_in)
        adjusted_sf = float(base_sf) + adjustment
        price = _taixing_ccl_price(adjusted_sf * parent["sf_factor"] * parent["price_factor"] / parent["opens"] * multiplier)
        total = _calc_total(quantity, price)
        adjust_text = f"+{adjustment:.6g}" if adjustment >= 0 else f"{adjustment:.6g}"
        formula = (
            f"({base_sf:.6g}{adjust_text})*{parent['sf_factor']:.6g}*"
            f"{parent['price_factor']:g}/{parent['opens']}*{multiplier:.6g}"
        )
        side_label = "双面" if side == "double" else "单面" if side == "single" else "同价"
        note = (
            f"命中泰兴CCL报价 Sheet {row.sheet} 第 {row.excel_row} 行，结构={row.stack}，"
            f"铜箔组={copper_group}，报价说明={family} {copper_group}{side_label}"
            f"{adjust_text}元/SF，父级{parent['parent_label']}，SF基础价={base_sf:.6g}，公式={formula}{structure_note}"
        )
        return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, f"{copper_group}:SF")

    percent = sheet_note.get("percents", {}).get(family)
    if percent is not None:
        regular_parent = _taixing_select_parent(length_in, width_in)
        if not regular_parent:
            return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "需人工确认：报价单说明未覆盖该情况")
        price_key = f"{copper_group}:{regular_parent['source_key']}"
        base_price = row.prices.get(price_key)
        if base_price is None:
            return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"找到CCL报价行，但{copper_group}组{regular_parent['source_label']}列无价格")
        multiplier = _taixing_ccl_multiplier(length_in, width_in)
        price = _taixing_ccl_price(float(base_price) * (1 + percent) * regular_parent["price_factor"] / regular_parent["opens"] * multiplier)
        total = _calc_total(quantity, price)
        formula = f"{base_price:.6g}*(1{percent:+.6g})*{regular_parent['price_factor']:g}/{regular_parent['opens']}*{multiplier:.6g}"
        note = (
            f"命中泰兴CCL报价 Sheet {row.sheet} 第 {row.excel_row} 行，结构={row.stack}，"
            f"铜箔组={copper_group}，报价说明={family}{percent:+.2%}，父级{regular_parent['parent_label']}，"
            f"尺寸列{regular_parent['source_label']}，公式={formula}{structure_note}"
        )
        return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, f"{copper_group}:{regular_parent['source_label']}")

    return ExtCalcResult("失败", "CCL", "待确认", "", "", "", "需人工确认：报价单说明未覆盖该情况")


def _taixing_sheet_adjustment(sheet_note: dict[str, Any], family: str, copper_group: str, side: str) -> tuple[str, float | None]:
    amounts = sheet_note.get("amounts", {})
    key = (family, copper_group, side)
    if key in amounts:
        return "explicit_sf_amount", float(amounts[key])
    if family in sheet_note.get("same_families", set()):
        return "same_price", None
    if family in sheet_note.get("base_families", set()):
        return "base_price", None
    return "", None


def _taixing_regular_ccl_result(
    row: ExtCclRule,
    copper_group: str,
    length_in: float,
    width_in: float,
    quantity: Any,
    structure_note: str,
) -> ExtCalcResult:
    parent = _taixing_select_parent(length_in, width_in)
    if not parent:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"无法识别泰兴CCL一开几父级：{length_in:g}X{width_in:g}")
    price_key = f"{copper_group}:{parent['source_key']}"
    base_price = row.prices.get(price_key)
    if base_price is None:
        return ExtCalcResult("失败", "CCL", "待确认", "", "", "", f"找到CCL报价行，但{copper_group}组{parent['source_label']}列无价格")
    multiplier = _taixing_ccl_multiplier(length_in, width_in)
    price = _taixing_ccl_price(float(base_price) * parent["price_factor"] / parent["opens"] * multiplier)
    total = _calc_total(quantity, price)
    formula = f"{base_price:.6g}*{parent['price_factor']:g}/{parent['opens']}*{multiplier:.6g}"
    note = (
        f"命中泰兴CCL报价 Sheet {row.sheet} 第 {row.excel_row} 行，结构={row.stack}，"
        f"铜箔组={copper_group}，父级{parent['parent_label']}，尺寸列{parent['source_label']}，"
        f"经向一开{parent['opens_w']}，纬向一开{parent['opens_h']}，总开数{parent['opens']}，"
        f"公式={formula}{structure_note}"
    )
    return ExtCalcResult("成功", "CCL", price, total, "", "", note, row.excel_row, f"{copper_group}:{parent['source_label']}")


def _taixing_select_parent(length_in: float, width_in: float) -> dict | None:
    candidates = [
        (37, 49, "36*48", '36"X48"', "37*49", 1),
        (41, 49, "40*48", '40"X48"', "41*49", 1),
        (43, 49, "42*48", '42"X48"', "43*49", 1),
        (74, 49, "36*48", '36"X48"', "37*49*2", 2),
        (82, 49, "40*48", '40"X48"', "41*49*2", 2),
        (86, 49, "42*48", '42"X48"', "43*49*2", 2),
    ]
    valid = []
    for parent_w, parent_h, source_key, source_label, parent_label, factor in candidates:
        opens_w = math.floor((parent_w + 1e-9) / length_in) if length_in else 0
        opens_h = math.floor((parent_h + 1e-9) / width_in) if width_in else 0
        opens = opens_w * opens_h
        if opens <= 0:
            continue
        fit_error = abs(length_in * opens_w - parent_w) + abs(width_in * opens_h - parent_h)
        valid.append(
            {
                "parent_w": parent_w,
                "parent_h": parent_h,
                "parent_label": parent_label,
                "source_key": source_key,
                "source_label": source_label,
                "price_factor": factor,
                "opens_w": opens_w,
                "opens_h": opens_h,
                "opens": opens,
                "fit_error": fit_error,
            }
        )
    if not valid:
        return None
    return sorted(valid, key=lambda item: (item["fit_error"], -item["parent_w"], -item["opens"]))[0]


def _taixing_ccl_sf_parent(length_in: float, width_in: float) -> dict | None:
    parent = _taixing_select_parent(length_in, width_in)
    if not parent:
        return None
    sf_factor_map = {
        "36*48": 12.0,
        "40*48": 13.33,
        "42*48": 13.33 * 1.05,
    }
    sf_factor = sf_factor_map.get(parent["source_key"])
    if sf_factor is None:
        return None
    return {**parent, "sf_factor": sf_factor}


def _taixing_ccl_multiplier(length_in: float, width_in: float) -> float:
    return 1.13


def _same_size(length_in: float, width_in: float, expected_length: float, expected_width: float) -> bool:
    return (abs(length_in - expected_length) <= 0.08 and abs(width_in - expected_width) <= 0.08) or (
        abs(width_in - expected_length) <= 0.08 and abs(length_in - expected_width) <= 0.08
    )


def _eaton_copper_and_foil(ws, row_idx: int) -> tuple[str, str]:
    copper = ""
    foil = ""
    for col in range(4, min(ws.max_column, 8) + 1):
        value = ws.cell(row_idx, col).value
        copper = copper or _norm_copper(value)
        foil = foil or _norm_foil(value)
    return copper, foil


def _looks_like_glass_pair(left: str, right: str) -> bool:
    return bool(re.fullmatch(r"\d{3,4}", left) and re.fullmatch(r"\d{3,4}", right))


def _find_header_row(ws) -> tuple[int | None, dict[str, int]]:
    for row_idx in range(1, min(ws.max_row, 30) + 1):
        headers = {_text(cell.value): cell.column for cell in ws[row_idx] if _text(cell.value)}
        if {"客户规格"} & set(headers) or {"物料长描述"} & set(headers) or {"物料描述"} & set(headers) or {"规格"} & set(headers):
            return row_idx, headers
    return None, {}


def _first_col(headers: dict[str, int], names: set[str]) -> int | None:
    return next((headers[name] for name in names if name in headers), None)


def _looks_like_pp(desc: str) -> bool:
    upper = desc.upper()
    return "RC" in upper and (
        re.search(r"\b\d+(?:\.\d+)?\s*M\b", upper) is not None
        or re.search(r"\d+(?:\.\d+)?\s*MM\s*[*xX×]\s*\d+(?:\.\d+)?\s*MM", upper) is not None
    )


def _extract_pp_small_piece_size(desc: str) -> tuple[float | None, float | None]:
    match = re.search(r"(\d+(?:\.\d+)?)\s*MM\s*[*xX×]\s*(\d+(?:\.\d+)?)\s*MM", desc, re.I)
    if not match:
        return None, None
    return float(match.group(1)) / 1000, float(match.group(2)) / 25.4


def _result_equal(actual: Any, expected: Any, *, tolerance: float) -> bool:
    a = _to_float(actual)
    e = _to_float(expected)
    return a is not None and e is not None and abs(a - e) <= tolerance


def _taixing_conflict_specs(ws, header_row: int, desc_col: int, expected_col: int) -> dict[str, str]:
    values: dict[str, list[tuple[int, str]]] = {}
    for row_idx in range(header_row + 1, ws.max_row + 1):
        spec = _text(ws.cell(row_idx, desc_col).value)
        price = _text(ws.cell(row_idx, expected_col).value)
        if not spec or not price:
            continue
        values.setdefault(spec, []).append((row_idx, price))
    conflicts: dict[str, str] = {}
    for spec, rows in values.items():
        nums = [_to_float(price) for _, price in rows]
        nums = [num for num in nums if num is not None]
        has_conflict = len(nums) >= 2 and max(nums) - min(nums) > 0.0002
        if has_conflict:
            conflicts[spec] = "；".join(f"第{row}行={price}" for row, price in rows)
    return conflicts


def _fmt_dim(value: float) -> str:
    value = round(float(value), 2)
    return str(int(value)) if value.is_integer() else f"{value:g}"


def _fmt_width(value: float | None) -> str:
    return f"{value:g}IN" if value else ""


def _fmt_length(value: int | None) -> str:
    return f"{value}m" if value else ""
