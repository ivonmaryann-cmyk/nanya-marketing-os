from __future__ import annotations

import json
import math
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename

from .db import append_job_log, create_job, get_job, update_job_status
from .job_control import launch_job_process
from .paths import JOBS_DIR


FEATURE = "order_reprice"
RULE_VERSION = "订单改价内置规则 v1"
EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")

MODE_LABELS = {
    "block1": "第一块：客户明细与430厂内明细匹配",
    "block2": "第二块：430价格核对",
    "block3": "第三块：客户改价结果与411厂内价格核对",
}


@dataclass
class MatchResult:
    status: str
    factory_indexes: list[int]
    factory_items: list[str]
    note: str = ""


@dataclass
class QuoteRow:
    source_file: str
    sheet_name: str
    excel_row: int
    product: str
    material_type: str
    values: dict[str, Any]
    size_prices: dict[str, float]


@dataclass
class ParsedSpec:
    material_type: str
    product: str | None
    thickness: float | None = None
    copper: str | None = None
    foil: str | None = None
    copper_state: str | None = None
    stack: str | None = None
    size_key: str | None = None
    glass: str | None = None
    rc: float | None = None
    length_m: float | None = None


def queue_order_reprice_job(
    employee_id: str,
    mode: str,
    customer_file,
    factory_file,
    quote_files: list,
) -> int:
    if mode not in MODE_LABELS:
        raise ValueError("未知订单改价处理块")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    employee_dir = JOBS_DIR / employee_id
    employee_dir.mkdir(parents=True, exist_ok=True)
    job_dir = employee_dir / f"{timestamp}_order_reprice_{mode}"
    job_dir.mkdir(parents=True, exist_ok=True)

    customer_path = _save_upload(customer_file, job_dir, "customer")
    factory_path = _save_upload(factory_file, job_dir, "factory")
    quote_paths = [_save_upload(file_obj, job_dir, f"quote_{idx}") for idx, file_obj in enumerate(quote_files, 1)]

    manifest = {
        "mode": mode,
        "label": MODE_LABELS[mode],
        "customer_path": str(customer_path),
        "factory_path": str(factory_path),
        "quote_paths": [str(path) for path in quote_paths],
    }
    manifest_path = job_dir / "order_reprice_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    source_filename = f"{MODE_LABELS[mode]}：{customer_file.filename} + {factory_file.filename}"
    if quote_paths:
        source_filename += f" + {len(quote_paths)}份报价单"
    job_id = create_job(employee_id, source_filename, str(manifest_path), RULE_VERSION, feature=FEATURE)
    launch_job_process(job_id, FEATURE, employee_id)
    return job_id


def run_order_reprice_job(job_id: int, employee_id: str) -> None:
    job = get_job(job_id)
    if not job or job["employee_id"] != employee_id:
        return

    update_job_status(job_id, status="running")
    try:
        manifest = json.loads(Path(job["stored_input_path"]).read_text(encoding="utf-8"))
        mode = manifest["mode"]
        append_job_log(job_id, f"开始处理订单改价任务：{MODE_LABELS.get(mode, mode)}")

        customer_path = Path(manifest["customer_path"])
        factory_path = Path(manifest["factory_path"])
        quote_paths = [Path(path) for path in manifest.get("quote_paths", [])]

        if mode == "block1":
            result = process_block1(customer_path, factory_path, job_id=job_id)
        elif mode == "block2":
            result = process_block2(customer_path, factory_path, quote_paths, job_id=job_id)
        elif mode == "block3":
            result = process_block3(customer_path, factory_path, job_id=job_id)
        else:
            raise ValueError(f"未知订单改价处理块：{mode}")

        output_path = Path(job["stored_input_path"]).with_name(
            f"订单改价结果_{mode}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        save_result_workbook(result["sheets"], output_path)
        update_job_status(
            job_id,
            status="completed",
            stored_result_path=str(output_path),
            success_count=result["success_count"],
            fail_count=result["fail_count"],
            skip_count=result.get("skip_count", 0),
            current_row=result["total_rows"],
            total_rows=result["total_rows"],
            completed=True,
        )
        append_job_log(
            job_id,
            f"订单改价任务完成：成功 {result['success_count']}，失败 {result['fail_count']}，结果已生成",
        )
    except Exception as exc:
        append_job_log(job_id, f"订单改价任务失败：{exc}")
        update_job_status(job_id, status="failed", error_message=str(exc), completed=True)
        raise


def process_block1(customer_path: Path, factory_path: Path, *, job_id: int | None = None) -> dict:
    customer_df = read_business_sheet(customer_path, {"采购订单号", "项次", "料件编号"})
    factory_df = read_business_sheet(factory_path, {"客户订单", "项次", "客户产品编号"})
    matches = match_customer_to_factory(
        customer_df,
        factory_df,
        customer_cols=("采购订单号", "项次", "料件编号"),
        factory_cols=("客户订单", "项次", "客户产品编号"),
    )
    return build_match_result(customer_df, factory_df, matches, job_id=job_id)


def process_block2(
    customer_path: Path,
    factory_path: Path,
    quote_paths: list[Path],
    *,
    job_id: int | None = None,
) -> dict:
    if not quote_paths:
        raise ValueError("第二块需要上传至少一份胜宏报价单")

    customer_df = read_business_sheet(customer_path, {"采购订单号", "项次", "料件编号", "规格", "采购量", "采购单位", "含税单价"})
    factory_df = read_business_sheet(factory_path, {"客户订单", "项次", "客户产品编号"})
    matches = match_customer_to_factory(
        customer_df,
        factory_df,
        customer_cols=("采购订单号", "项次", "料件编号"),
        factory_cols=("客户订单", "项次", "客户产品编号"),
    )
    quote_rows = load_quote_rows(quote_paths)
    append_job_log(job_id, f"已读取胜宏报价单：{len(quote_rows)} 条报价记录", total_rows=len(customer_df)) if job_id else None

    output = customer_df.copy()
    notes: list[dict[str, Any]] = []
    success = 0
    fail = 0
    actual_qty_values = []
    factory_items = []
    quote_prices = []
    compare_results = []
    quote_notes = []

    for pos, (idx, row) in enumerate(customer_df.iterrows(), 1):
        spec = _text(row.get("规格"))
        actual_qty = calculate_actual_quantity(row.get("采购单位"), row.get("采购量"), spec)
        actual_qty_values.append(actual_qty)

        match = matches.get(idx, MatchResult("未匹配", [], [], ""))
        factory_items.append(",".join(match.factory_items))

        quote = find_quote_price(spec, quote_rows)
        quote_prices.append(quote["price"])
        quote_notes.append(quote["note"])

        if quote["price"] is None:
            compare = "未命中报价"
            fail += 1
            notes.append(
                {
                    "行号": pos + 1,
                    "采购订单号": row.get("采购订单号"),
                    "项次": row.get("项次"),
                    "料件编号": row.get("料件编号"),
                    "规格": spec,
                    "原因": quote["note"],
                }
            )
        elif _prices_equal(quote["price"], row.get("含税单价")):
            compare = "正确"
            success += 1
        else:
            compare = "错误"
            fail += 1
        compare_results.append(compare)

        if job_id and (pos == 1 or pos % 50 == 0 or pos == len(customer_df)):
            append_job_log(
                job_id,
                f"第二块价格核对进度：{pos}/{len(customer_df)}",
                success_count=success,
                fail_count=fail,
                current_row=pos,
                total_rows=len(customer_df),
            )

    output["实际订单数量"] = actual_qty_values
    output["厂内项次"] = factory_items
    output["报价单价格"] = quote_prices
    output["价格比对结果"] = compare_results
    output["报价命中说明"] = quote_notes

    summary = summary_sheet(
        [
            ("总记录数", len(output)),
            ("价格正确数量", compare_results.count("正确")),
            ("价格错误数量", compare_results.count("错误")),
            ("未命中报价数量", compare_results.count("未命中报价")),
            ("未匹配厂内数量", sum(1 for item in matches.values() if item.status == "未匹配")),
        ]
    )
    return {
        "sheets": [
            ("客户价格核对结果", output),
            ("匹配汇总", summary),
            ("报价未命中说明", pd.DataFrame(notes)),
        ],
        "success_count": success,
        "fail_count": fail,
        "skip_count": 0,
        "total_rows": len(output),
    }


def process_block3(customer_path: Path, factory_path: Path, *, job_id: int | None = None) -> dict:
    customer_df = read_business_sheet(customer_path, {"采购订单号", "项次", "料件编号", "规格", "含税单价"})
    factory_df = read_business_sheet(factory_path, {"客户单号", "项次", "客户产品编号", "单价"})
    matches = match_customer_to_factory(
        customer_df,
        factory_df,
        customer_cols=("采购订单号", "项次", "料件编号"),
        factory_cols=("客户单号", "项次", "客户产品编号"),
    )

    output = customer_df.copy()
    statuses = []
    factory_items = []
    factory_prices = []
    factory_price_notes = []
    compare_results = []
    success = 0
    fail = 0

    for pos, (idx, row) in enumerate(customer_df.iterrows(), 1):
        match = matches.get(idx, MatchResult("未匹配", [], [], ""))
        statuses.append(match.status)
        factory_items.append(",".join(match.factory_items))
        price_parts = [
            _block3_factory_price(factory_df.loc[fidx], row.get("规格"))
            for fidx in _sort_factory_indexes_by_item(factory_df, match.factory_indexes)
        ]
        matched_price, price_note = _pick_block3_price(price_parts, len(match.factory_indexes) > 1)
        factory_price_notes.append(price_note)
        factory_prices.append(matched_price)
        if matched_price is None:
            result = "未匹配"
            fail += 1
        elif _prices_equal(matched_price, row.get("含税单价")):
            result = "正确"
            success += 1
        else:
            result = "错误"
            fail += 1
        compare_results.append(result)

        if job_id and (pos == 1 or pos % 50 == 0 or pos == len(customer_df)):
            append_job_log(
                job_id,
                f"第三块价格核对进度：{pos}/{len(customer_df)}",
                success_count=success,
                fail_count=fail,
                current_row=pos,
                total_rows=len(customer_df),
            )

    output["匹配状态"] = statuses
    output["厂内项次"] = factory_items
    output["厂内匹配价格"] = factory_prices
    output["厂内价格换算说明"] = factory_price_notes
    output["价格核对结果"] = compare_results
    summary = summary_sheet(
        [
            ("总记录数", len(output)),
            ("价格正确数量", compare_results.count("正确")),
            ("价格错误数量", compare_results.count("错误")),
            ("未匹配数量", compare_results.count("未匹配")),
        ]
    )
    return {
        "sheets": [("客户改价核对结果", output), ("核对汇总", summary)],
        "success_count": success,
        "fail_count": fail,
        "skip_count": 0,
        "total_rows": len(output),
    }


def build_match_result(
    customer_df: pd.DataFrame,
    factory_df: pd.DataFrame,
    matches: dict[int, MatchResult],
    *,
    job_id: int | None = None,
) -> dict:
    customer_result = customer_df.copy()
    customer_statuses = []
    customer_factory_items = []
    factory_to_customer: dict[int, list[str]] = {}
    success = 0
    fail = 0

    for pos, (idx, _) in enumerate(customer_df.iterrows(), 1):
        match = matches.get(idx, MatchResult("未匹配", [], [], ""))
        customer_statuses.append(match.status)
        customer_factory_items.append(",".join(match.factory_items))
        if match.status in {"已匹配", "拆分匹配"}:
            success += 1
        else:
            fail += 1
        for fidx in match.factory_indexes:
            factory_to_customer.setdefault(fidx, []).append(_text(customer_df.loc[idx].get("项次")))
        if job_id and (pos == 1 or pos % 50 == 0 or pos == len(customer_df)):
            append_job_log(
                job_id,
                f"匹配进度：{pos}/{len(customer_df)}",
                success_count=success,
                fail_count=fail,
                current_row=pos,
                total_rows=len(customer_df),
            )

    customer_result["匹配状态"] = customer_statuses
    customer_result["厂内项次"] = customer_factory_items

    factory_result = factory_df.copy()
    factory_statuses = []
    factory_customer_items = []
    for idx in factory_df.index:
        linked = factory_to_customer.get(idx, [])
        factory_statuses.append("已匹配" if linked else "未匹配")
        factory_customer_items.append(",".join(linked))
    factory_result["匹配状态"] = factory_statuses
    factory_result["客户项次"] = factory_customer_items

    summary = summary_sheet(
        [
            ("客户总记录数", len(customer_result)),
            ("客户已匹配数量", success),
            ("客户未匹配数量", fail),
            ("厂内总记录数", len(factory_result)),
            ("厂内已匹配数量", factory_statuses.count("已匹配")),
            ("厂内未匹配数量", factory_statuses.count("未匹配")),
        ]
    )
    return {
        "sheets": [
            ("客户明细匹配结果", customer_result),
            ("厂内明细匹配结果", factory_result),
            ("匹配汇总", summary),
        ],
        "success_count": success,
        "fail_count": fail,
        "skip_count": 0,
        "total_rows": len(customer_result),
    }


def match_customer_to_factory(
    customer_df: pd.DataFrame,
    factory_df: pd.DataFrame,
    *,
    customer_cols: tuple[str, str, str],
    factory_cols: tuple[str, str, str],
) -> dict[int, MatchResult]:
    c_po, c_item, c_mat = customer_cols
    f_po, f_item, f_mat = factory_cols

    factory_exact: dict[tuple[str, str, str], list[int]] = {}
    factory_by_order_mat: dict[tuple[str, str], list[int]] = {}
    for fidx, row in factory_df.iterrows():
        key = (_norm(row.get(f_po)), _norm_item(row.get(f_item)), _norm(row.get(f_mat)))
        if not all(key):
            continue
        factory_exact.setdefault(key, []).append(fidx)
        factory_by_order_mat.setdefault((key[0], key[2]), []).append(fidx)

    customer_items_by_order_mat: dict[tuple[str, str], set[str]] = {}
    for _, row in customer_df.iterrows():
        po = _norm(row.get(c_po))
        item = _norm_item(row.get(c_item))
        mat = _norm(row.get(c_mat))
        if po and item and mat:
            customer_items_by_order_mat.setdefault((po, mat), set()).add(item)

    results: dict[int, MatchResult] = {}
    for cidx, row in customer_df.iterrows():
        po = _norm(row.get(c_po))
        item = _norm_item(row.get(c_item))
        mat = _norm(row.get(c_mat))
        if not po or not item or not mat:
            results[cidx] = MatchResult("未匹配", [], [], "关键字段为空")
            continue

        exact = factory_exact.get((po, item, mat), [])
        if exact:
            results[cidx] = MatchResult("已匹配", exact, [_norm_item(factory_df.loc[idx].get(f_item)) for idx in exact])
            continue

        candidates = []
        ambiguous_items = []
        for fidx in factory_by_order_mat.get((po, mat), []):
            factory_item = _norm_item(factory_df.loc[fidx].get(f_item))
            if not factory_item.startswith(item) or factory_item == item:
                continue
            possible_customer_items = [
                candidate
                for candidate in customer_items_by_order_mat.get((po, mat), set())
                if factory_item.startswith(candidate) and factory_item != candidate
            ]
            if len(possible_customer_items) > 1:
                ambiguous_items.append(factory_item)
            else:
                candidates.append(fidx)
        if ambiguous_items:
            results[cidx] = MatchResult("拆分项次歧义", [], [], f"厂内项次可能对应多个客户项次：{','.join(ambiguous_items)}")
        elif candidates:
            results[cidx] = MatchResult(
                "拆分匹配",
                candidates,
                [_norm_item(factory_df.loc[idx].get(f_item)) for idx in candidates],
            )
        else:
            results[cidx] = MatchResult("未匹配", [], [], "未找到对应厂内记录")
    return results


def load_quote_rows(paths: list[Path]) -> list[QuoteRow]:
    rows: list[QuoteRow] = []
    for path in paths:
        workbook = pd.ExcelFile(path)
        for sheet_name in workbook.sheet_names:
            raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
            for header_idx in range(len(raw)):
                headers = [_clean_header(value) for value in raw.iloc[header_idx].tolist()]
                if _is_ccl_header(headers):
                    rows.extend(_parse_ccl_rows(raw, headers, header_idx, path.name, sheet_name))
                elif _is_pp_header(headers):
                    rows.extend(_parse_pp_rows(raw, headers, header_idx, path.name, sheet_name))
    return rows


def _parse_ccl_rows(raw, headers: list[str], header_idx: int, source_file: str, sheet_name: str) -> list[QuoteRow]:
    product_col = _find_col(headers, {"产品型号", "产品名称"})
    thickness_col = _find_col(headers, {"厚度mm", "厚度（mm）", "厚度"})
    copper_col = _find_col(headers, {"铜箔", "铜箔厚度"})
    foil_col = _find_col(headers, {"铜箔特性", "铜箔类型"})
    state_col = _find_col(headers, {"板材类型", "是否含铜"})
    stack_col = _find_col(headers, {"叠构", "组合结构"})
    sf_col = _find_col(headers, {"SF单价", "SF价格"})
    size_cols = {headers[idx].upper().replace("X", "*"): idx for idx, value in enumerate(headers) if re.fullmatch(r"\d+(?:\.\d+)?[*X]\d+(?:\.\d+)?", value.upper())}
    if product_col is None or thickness_col is None or sf_col is None:
        return []

    rows = []
    for row_idx in range(header_idx + 1, len(raw)):
        values = raw.iloc[row_idx].tolist()
        if _looks_like_header(values) or _row_is_empty(values):
            continue
        product = _norm_product(values[product_col] if product_col < len(values) else "")
        if not product:
            continue
        thickness_value = values[thickness_col] if thickness_col is not None and thickness_col < len(values) else None
        row_values = {
            "厚度": _number(thickness_value),
            "厚度候选": _numbers(thickness_value),
            "铜箔": _norm_copper(values[copper_col] if copper_col is not None and copper_col < len(values) else None),
            "铜箔特性": _norm_foil(values[foil_col] if foil_col is not None and foil_col < len(values) else None),
            "板材类型": _norm_state(values[state_col] if state_col is not None and state_col < len(values) else None),
            "叠构": _norm_stack(values[stack_col] if stack_col is not None and stack_col < len(values) else None),
            "SF单价": _number(values[sf_col] if sf_col is not None and sf_col < len(values) else None),
        }
        price_map = {}
        for size_key, col_idx in size_cols.items():
            price = _number(values[col_idx] if col_idx < len(values) else None)
            if price is not None:
                price_map[_norm_size_key(size_key)] = price
        if row_values["SF单价"] is None and not price_map:
            continue
        rows.append(QuoteRow(source_file, sheet_name, row_idx + 1, product, "CCL", row_values, price_map))
    return rows


def _parse_pp_rows(raw, headers: list[str], header_idx: int, source_file: str, sheet_name: str) -> list[QuoteRow]:
    product_col = _find_col(headers, {"产品型号", "产品名称"})
    glass_col = _find_col(headers, {"布种"})
    rc_col = _find_col(headers, {"RC值", "树脂含量"})
    length_col = _find_col(headers, {"长度M", "每卷长度"})
    price_col = _find_col(headers, {"RL价格", "每卷价格"})
    if price_col is None:
        price_col = _find_col(headers, {"单价M", "单价/M"})
    if product_col is None or glass_col is None or rc_col is None or price_col is None:
        return []

    rows = []
    for row_idx in range(header_idx + 1, len(raw)):
        values = raw.iloc[row_idx].tolist()
        if _looks_like_header(values) or _row_is_empty(values):
            continue
        product = _norm_product(values[product_col] if product_col < len(values) else "")
        glass = _norm(values[glass_col] if glass_col < len(values) else "")
        price = _number(values[price_col] if price_col < len(values) else None)
        if not product or not glass or price is None:
            continue
        rows.append(
            QuoteRow(
                source_file,
                sheet_name,
                row_idx + 1,
                product,
                "PP",
                {
                    "布种": _norm_glass(glass),
                    "RC": values[rc_col] if rc_col < len(values) else None,
                    "长度": _number(values[length_col] if length_col is not None and length_col < len(values) else None),
                    "价格": price,
                },
                {},
            )
        )
    return rows


def find_quote_price(spec: str, quote_rows: list[QuoteRow]) -> dict[str, Any]:
    parsed = parse_spec(spec)
    if parsed.product is None:
        return {"price": None, "note": "无法从规格提取产品型号"}
    product_aliases = _quote_product_aliases(parsed.product, parsed.material_type)
    candidates = [
        row
        for row in quote_rows
        if row.material_type == parsed.material_type and _norm_product(row.product) in product_aliases
    ]
    candidates.sort(key=lambda row: product_aliases.index(_norm_product(row.product)))
    if not candidates:
        note = f"报价单未找到产品型号 {parsed.product}"
        if len(product_aliases) > 1:
            note += f"（已尝试 {', '.join(product_aliases[1:])}）"
        return {"price": None, "note": note}
    if parsed.material_type == "PP":
        return _find_pp_quote(parsed, candidates)
    return _find_ccl_quote(parsed, candidates)


def _find_pp_quote(parsed: ParsedSpec, candidates: list[QuoteRow]) -> dict[str, Any]:
    for row in candidates:
        if parsed.glass and not _token_matches(row.values.get("布种"), parsed.glass):
            continue
        if parsed.rc is not None and not _rc_matches(row.values.get("RC"), parsed.rc):
            continue
        if parsed.length_m is not None and row.values.get("长度") is not None and not _float_equal(parsed.length_m, row.values["长度"]):
            continue
        return {
            "price": row.values.get("价格"),
            "note": f"命中 {row.source_file}/{row.sheet_name} 第{row.excel_row}行",
        }
    return {"price": None, "note": "未命中PP报价：产品、布种、RC或长度不匹配"}


def _find_ccl_quote(parsed: ParsedSpec, candidates: list[QuoteRow]) -> dict[str, Any]:
    for row in candidates:
        if parsed.thickness is not None and not _number_choice_matches(row.values.get("厚度候选") or row.values.get("厚度"), parsed.thickness, 0.003):
            continue
        if parsed.copper and row.values.get("铜箔") and parsed.copper != row.values["铜箔"]:
            continue
        if parsed.foil and row.values.get("铜箔特性") and parsed.foil != row.values["铜箔特性"]:
            continue
        if parsed.copper_state and row.values.get("板材类型") and parsed.copper_state != row.values["板材类型"]:
            continue
        if parsed.stack and row.values.get("叠构") and parsed.stack != row.values["叠构"]:
            continue
        price = row.size_prices.get(parsed.size_key or "") if parsed.size_key else None
        if price is None and row.values.get("SF单价") is not None:
            price = row.values["SF单价"]
        if price is not None:
            return {
                "price": price,
                "note": f"命中 {row.source_file}/{row.sheet_name} 第{row.excel_row}行",
            }
    return {"price": None, "note": "未命中CCL报价：型号、厚度、铜箔、叠构或尺寸不匹配"}


def parse_spec(spec: str) -> ParsedSpec:
    text = _text(spec).upper().replace("×", "*").replace("Ｘ", "*")
    product_match = re.search(r"(NY[-A-Z0-9()]+)\s*:", text)
    product = product_match.group(1) if product_match else None
    if not product:
        product_match = re.search(r"\b(NY[-A-Z0-9()]+P?)\b", text)
        product = product_match.group(1) if product_match else None
    material_type = "PP" if product and product.endswith("P") else "CCL"
    if material_type == "PP":
        glass_match = re.search(r":\s*([0-9]{2,4})\s*RC", text)
        rc_match = re.search(r"RC\s*=?\s*(\d+(?:\.\d+)?)\s*%?", text)
        length_match = re.search(r"(\d+(?:\.\d+)?)\s*M\s*/?\s*R", text)
        return ParsedSpec(
            material_type="PP",
            product=product,
            glass=_norm_glass(glass_match.group(1)) if glass_match else None,
            rc=float(rc_match.group(1)) if rc_match else None,
            length_m=float(length_match.group(1)) if length_match else None,
        )

    thickness_match = re.search(r"(\d+(?:\.\d+)?)\s*MM", text)
    copper_match = re.search(r"MM\s*([H\d.]+/[H\d.]+)", text)
    foil_match = re.search(r"\(([^)）]+)\)", text)
    size_match = re.search(r"(\d+(?:\.\d+)?)\s*[*X]\s*(\d+(?:\.\d+)?)", text)
    stack_match = re.search(r"\(([^()）]*\d{2,4}[^()）]*)\)\s*$", text)
    return ParsedSpec(
        material_type="CCL",
        product=product,
        thickness=float(thickness_match.group(1)) if thickness_match else None,
        copper=_norm_copper(copper_match.group(1)) if copper_match else None,
        foil=_norm_foil(foil_match.group(1)) if foil_match else None,
        copper_state="含铜" if "含铜" in text and "不含铜" not in text else "不含铜" if "不含铜" in text else None,
        stack=_norm_stack(stack_match.group(1)) if stack_match else None,
        size_key=_norm_size_key(f"{size_match.group(1)}*{size_match.group(2)}") if size_match else None,
    )


def calculate_actual_quantity(unit: Any, qty: Any, spec: str) -> float | None:
    quantity = _number(qty)
    if quantity is None:
        return None
    if _text(unit) == "卷":
        parsed = parse_spec(spec)
        if parsed.length_m is not None:
            return parsed.length_m * quantity
    return quantity


def _block3_factory_price(factory_row: pd.Series, customer_spec: Any) -> tuple[float | None, str, str]:
    factory_item = _norm_item(factory_row.get("项次"))
    unit_price = _number(factory_row.get("单价"))
    if unit_price is None:
        return None, "厂内单价为空", factory_item

    specs = [
        customer_spec,
        factory_row.get("规格"),
        factory_row.get("客户规格"),
        factory_row.get("品名"),
    ]
    if not any(_is_pp_spec(spec) for spec in specs):
        base_price = _block3_non_pp_price(factory_row)
        if base_price is None:
            return None, "非PP: 单价为空", factory_item
        return base_price, f"非PP: 单价{base_price:g}", factory_item

    length = _pp_length_from_specs(specs)
    if length is None:
        return unit_price, f"PP长度缺失，未换算: 单价{unit_price:g}", factory_item

    converted = unit_price * length
    return converted, f"PP: 单价{unit_price:g}×{length:g}={converted:g}", factory_item


def _pick_block3_price(price_parts: list[tuple[float | None, str, str]], is_split: bool) -> tuple[float | None, str]:
    notes = []
    for price, note, factory_item in price_parts:
        if price is not None:
            if is_split:
                prefix = f"拆分取用项次{factory_item}，未累加"
                return price, f"{prefix}; {note}" if note else prefix
            return price, note
        if note:
            notes.append(note)
    return None, "; ".join(notes)


def _sort_factory_indexes_by_item(factory_df: pd.DataFrame, factory_indexes: list[int]) -> list[int]:
    return sorted(factory_indexes, key=lambda idx: _item_sort_key(factory_df.loc[idx].get("项次")))


def _item_sort_key(value: Any) -> tuple[int, str]:
    text = _norm_item(value)
    number = _number(text)
    return (int(number) if number is not None else 10**9, text)


def _block3_non_pp_price(factory_row: pd.Series) -> float | None:
    return _number(factory_row.get("单价"))


def _is_pp_spec(spec: Any) -> bool:
    text = _text(spec).upper().replace("×", "*").replace("Ｘ", "*")
    if not text:
        return False
    product_match = re.search(r"\b(NY[-A-Z0-9()]+)\b", text)
    if product_match:
        product = _norm_product(product_match.group(1))
        if _product_is_pp(product):
            return True
    return "RC" in text and _pp_length_from_spec(text) is not None


def _product_is_pp(product: str) -> bool:
    normalized = _norm_product(product)
    core = re.sub(r"\([^)]*\)$", "", normalized)
    return core.endswith("P")


def _pp_length_from_specs(specs: list[Any]) -> float | None:
    for spec in specs:
        length = _pp_length_from_spec(spec)
        if length is not None:
            return length
    return None


def _pp_length_from_spec(spec: Any) -> float | None:
    text = _text(spec).upper().replace("×", "*").replace("Ｘ", "*")
    if not text:
        return None
    parsed = parse_spec(text)
    if parsed.length_m is not None:
        return parsed.length_m
    patterns = [
        r"(\d+(?:\.\d+)?)\s*M\s*/?\s*卷",
        r"卷\s*(\d+(?:\.\d+)?)\s*M",
        r"(\d+(?:\.\d+)?)\s*M\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return float(match.group(1))
    return None


def read_business_sheet(path: Path, required_columns: set[str]) -> pd.DataFrame:
    workbook = pd.ExcelFile(path)
    for sheet_name in workbook.sheet_names:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
        header_idx = find_header_row(raw, required_columns)
        if header_idx is None:
            continue
        df = pd.read_excel(path, sheet_name=sheet_name, header=header_idx, dtype=object)
        df = df.dropna(how="all").reset_index(drop=True)
        df.columns = [str(col).strip() for col in df.columns]
        return df
    raise ValueError(f"{path.name} 未找到包含字段 {', '.join(sorted(required_columns))} 的业务 Sheet")


def find_header_row(raw: pd.DataFrame, required_columns: set[str]) -> int | None:
    required = {_clean_header(col) for col in required_columns}
    for idx in range(min(20, len(raw))):
        headers = {_clean_header(value) for value in raw.iloc[idx].tolist() if _text(value)}
        if required.issubset(headers):
            return idx
    return None


def save_result_workbook(sheets: list[tuple[str, pd.DataFrame]], output_path: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    header_fill = PatternFill("solid", fgColor="E8F0FE")
    for sheet_name, df in sheets:
        ws = wb.create_sheet(_safe_sheet_name(sheet_name))
        if df is None or df.empty:
            ws.append(["说明"])
            ws.append(["无数据"])
        else:
            headers = [str(col) for col in df.columns]
            ws.append(headers)
            for row in df.itertuples(index=False):
                ws.append([_excel_value(value) for value in row])
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
        ws.freeze_panes = "A2"
        for col_idx, column in enumerate(ws.columns, 1):
            values = [str(cell.value) for cell in list(column)[:80] if cell.value is not None]
            width = min(max([len(value) for value in values] + [10]) + 2, 36)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
    wb.save(output_path)


def summary_sheet(items: list[tuple[str, Any]]) -> pd.DataFrame:
    return pd.DataFrame(items, columns=["项目", "数量"])


def _save_upload(file_obj, target_dir: Path, prefix: str) -> Path:
    filename = file_obj.filename or f"{prefix}.xlsx"
    if not filename.lower().endswith(EXCEL_EXTENSIONS):
        raise ValueError(f"{filename} 不是支持的 Excel 文件")
    safe_name = secure_filename(filename) or f"{prefix}.xlsx"
    path = target_dir / f"{prefix}_{safe_name}"
    file_obj.save(path)
    return path


def _is_ccl_header(headers: list[str]) -> bool:
    joined = "|".join(headers)
    return ("产品型号" in joined or "产品名称" in joined) and "厚度" in joined and ("叠构" in joined or "组合结构" in joined)


def _is_pp_header(headers: list[str]) -> bool:
    joined = "|".join(headers)
    return ("产品型号" in joined or "产品名称" in joined) and "布种" in joined and ("RC值" in joined or "树脂含量" in joined)


def _find_col(headers: list[str], aliases: set[str]) -> int | None:
    clean_aliases = {_clean_header(alias) for alias in aliases}
    for idx, header in enumerate(headers):
        if header in clean_aliases:
            return idx
    return None


def _looks_like_header(values: list[Any]) -> bool:
    text = "|".join(_clean_header(value) for value in values)
    return "产品型号" in text or "产品名称" in text


def _row_is_empty(values: list[Any]) -> bool:
    return not any(_text(value) for value in values)


def _prices_equal(left: Any, right: Any) -> bool:
    left_num = _number(left)
    right_num = _number(right)
    if left_num is None or right_num is None:
        return False
    return round(left_num + 1e-9, 2) == round(right_num + 1e-9, 2)


def _quote_product_aliases(product: str, material_type: str) -> list[str]:
    primary = _norm_product(product)
    aliases = [primary]
    if material_type == "PP" and primary.endswith("P") and len(primary) > 1:
        aliases.append(primary[:-1])
    return list(dict.fromkeys(alias for alias in aliases if alias))


def _token_matches(rule_value: Any, target: str) -> bool:
    normalized_target = _norm_glass(target)
    if not normalized_target:
        return True
    return normalized_target in _token_alternatives(rule_value, _norm_glass)


def _token_alternatives(value: Any, normalizer) -> list[str]:
    text = _text(value)
    if not text:
        return []
    return [
        normalizer(part)
        for part in re.split(r"[/,，、\s]+", text)
        if normalizer(part)
    ]


def _rc_matches(rule_value: Any, target: float) -> bool:
    text = _text(rule_value).replace("％", "%")
    if not text:
        return False
    numbers = [float(item) for item in re.findall(r"\d+(?:\.\d+)?", text)]
    if not numbers:
        return False
    if len(numbers) >= 2 and ("-" in text or "~" in text or "～" in text):
        low = _to_percent_value(numbers[0])
        high = _to_percent_value(numbers[1])
        return min(low, high) <= target <= max(low, high)
    value = _to_percent_value(numbers[0])
    return abs(value - target) <= 0.01


def _to_percent_value(value: float) -> float:
    return value * 100 if abs(value) <= 1 else value


def _float_equal(left: float | None, right: float | None, tolerance: float = 0.01) -> bool:
    return left is not None and right is not None and abs(left - right) <= tolerance


def _number_choice_matches(rule_value: Any, target: float, tolerance: float = 0.01) -> bool:
    numbers = rule_value if isinstance(rule_value, list) else _numbers(rule_value)
    if not numbers:
        return True
    return any(abs(target - number) <= tolerance for number in numbers)


def _number(value: Any) -> float | None:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _text(value).replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else None


def _numbers(value: Any) -> list[float]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return []
    if isinstance(value, (int, float)):
        return [float(value)]
    text = _text(value).replace(",", "")
    return [float(item) for item in re.findall(r"-?\d+(?:\.\d+)?", text)]


def _text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def _norm(value: Any) -> str:
    return re.sub(r"\s+", "", _text(value)).upper()


def _norm_item(value: Any) -> str:
    text = _norm(value)
    return text[:-2] if re.fullmatch(r"\d+\.0", text) else text


def _norm_product(value: Any) -> str:
    return _norm(value).replace("（", "(").replace("）", ")")


def _norm_glass(value: Any) -> str:
    return _norm(value)


def _norm_copper(value: Any) -> str:
    text = _norm(value).replace("T", "1").replace(" ", "")
    return text


def _norm_foil(value: Any) -> str:
    text = _norm(value)
    if "/" in text:
        text = text.split("/")[0]
    return text


def _norm_state(value: Any) -> str:
    text = _text(value)
    if "不含铜" in text:
        return "不含铜"
    if "含铜" in text:
        return "含铜"
    return _norm(value)


def _norm_size_key(value: Any) -> str:
    numbers = [_number(part) for part in re.split(r"[*Xx]", _text(value).upper())]
    if len(numbers) < 2 or numbers[0] is None or numbers[1] is None:
        return _norm(value).replace("X", "*")
    return f"{_format_size(numbers[0])}*{_format_size(numbers[1])}"


def _format_size(value: float) -> str:
    rounded = round(value)
    return str(int(rounded)) if abs(value - rounded) < 0.51 else f"{value:g}"


def _norm_stack(value: Any) -> str:
    text = _norm(value).replace("X", "*").replace("×", "*")
    parts: dict[str, int] = {}
    for token in re.split(r"[+/,，、]", text):
        token = token.strip()
        if not token:
            continue
        match = re.fullmatch(r"(?:(\d+)\*)?(\d{2,4})", token)
        if match:
            count = int(match.group(1)) if match.group(1) else 1
            glass = match.group(2)
        else:
            match = re.fullmatch(r"(\d{2,4})\*(\d+)", token)
            if not match:
                continue
            glass = match.group(1)
            count = int(match.group(2))
        parts[glass] = parts.get(glass, 0) + count
    if not parts:
        return text
    return "+".join(f"{glass}*{parts[glass]}" for glass in sorted(parts))


def _clean_header(value: Any) -> str:
    return re.sub(r"\s+", "", _text(value)).replace("（", "(").replace("）", ")")


def _excel_value(value: Any) -> Any:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if hasattr(value, "to_pydatetime"):
        return value.to_pydatetime()
    if hasattr(value, "to_pytimedelta"):
        return str(value)
    return value


def _safe_sheet_name(value: str) -> str:
    return re.sub(r"[\[\]\:\*\?\/\\]", "_", value)[:31]
