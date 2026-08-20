from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from threading import Lock
from typing import Any, Callable

from werkzeug.utils import secure_filename

from .ai_repair_config import AiRepairConfig, config_from_manifest_snapshot, get_ai_repair_config
from .db import append_job_log, create_job, get_job, update_job_status
from .docling_parser import parse_pdf_with_docling
from .image_table_parser import parse_image_tables
from .job_control import launch_job_process
from .json_to_excel import build_business_document, write_conversion_workbook
from .paths import JOBS_DIR
from .pdf_native_parser import parse_pdf_native
from .purchase_ai_repair import audit_and_repair_purchase_document
from .purchase_excel_writer import write_internal_sales_workbook, write_purchase_order_workbook
from .purchase_factory_mapper import project_factory_document, safe_result_stem
from .purchase_order_pipeline import run_purchase_order_pipeline
from .purchase_performance import add_stage_ms, file_sha256, pipeline_fingerprint
from .purchase_result_normalizer import normalize_purchase_document
from .template_parser import identify_template, likely_order_number, likely_supplier


FEATURE = "pdf_excel"
RULE_VERSION = "pdf_excel_v1"
ALLOWED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}


def _safe_input_filename(original_filename: str, index: int) -> str:
    path = Path(original_filename or "")
    suffix = path.suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        suffix = ".dat"
    stem = path.name[: -len(path.suffix)] if path.suffix else path.name
    safe_stem = secure_filename(stem).strip("._-") or f"file_{index:03d}"
    return f"{index:03d}_{safe_stem}{suffix}"


def queue_pdf_excel_job(
    employee_id: str,
    uploaded_files,
    *,
    max_workers: int = 2,
) -> int:
    files = [file_obj for file_obj in uploaded_files if file_obj and file_obj.filename]
    if not files:
        raise ValueError("请至少上传一个 PDF 或图片文件。")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    employee_dir = JOBS_DIR / employee_id
    job_dir = employee_dir / f"{timestamp}_pdf_excel"
    input_dir = job_dir / "inputs"
    input_dir.mkdir(parents=True, exist_ok=True)

    manifest_files: list[dict[str, str]] = []
    for index, file_obj in enumerate(files, start=1):
        original_filename = str(file_obj.filename or "").strip()
        suffix = Path(original_filename).suffix.lower()
        if suffix not in ALLOWED_EXTENSIONS:
            raise ValueError(f"不支持的文件类型：{original_filename}")
        stored_path = input_dir / _safe_input_filename(original_filename, index)
        file_obj.save(stored_path)
        hash_started = time.perf_counter()
        content_sha256 = file_sha256(stored_path)
        hash_ms = (time.perf_counter() - hash_started) * 1000
        manifest_files.append(
            {
                "original_filename": original_filename,
                "stored_path": str(stored_path),
                "extension": suffix,
                "content_sha256": content_sha256,
                "hash_ms": round(hash_ms, 3),
            }
        )

    ai_config = get_ai_repair_config()
    manifest_path = job_dir / "manifest.json"
    manifest = {
        "feature": FEATURE,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "pipeline_fingerprint": pipeline_fingerprint(),
        "max_workers": max(1, min(int(max_workers or 2), 4)),
        "ai_config": ai_config.safe_metadata(),
        "files": manifest_files,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    display_name = (
        f"PDF/图片转Excel：{manifest_files[0]['original_filename']}"
        if len(manifest_files) == 1
        else f"PDF/图片转Excel：{len(manifest_files)} 个文件"
    )
    job_id = create_job(employee_id, display_name, str(manifest_path), RULE_VERSION, feature=FEATURE)
    launch_job_process(job_id, FEATURE, employee_id)
    return job_id


def _document_base(source_file: str, file_type: str) -> dict[str, Any]:
    return {
        "source_file": source_file,
        "file_type": file_type,
        "parser_mode": "",
        "template_id": "",
        "template_label": "",
        "page_count": 0,
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "warnings": [],
        "order_number": "",
        "supplier": "",
        "pages": [],
        "docling_tables": [],
        "docling_classified_tables": [],
        "source_tables": [],
        "docling_lines": [],
    }


def _table_text(rows: list[list[str]]) -> str:
    return "\n".join(" ".join(str(value or "") for value in row) for row in rows)


def _docling_compact_text(table: dict[str, Any]) -> str:
    return "".join((str(table.get("title") or "") + "\n" + _table_text(table.get("rows") or [])).split()).lower()


def _has_any(text: str, keywords: list[str]) -> bool:
    return any(keyword.lower().replace(" ", "") in text for keyword in keywords)


def _docling_detail_score(table: dict[str, Any]) -> tuple[int, bool, int]:
    rows = table.get("rows") or []
    max_cols = max((len(row) for row in rows), default=0)
    text = _docling_compact_text(table)
    categories = {
        "seq": ["序号", "no.", "itemno", "item"],
        "identity": ["物料", "原料", "编码", "料号", "partno", "partno.", "description", "品号"],
        "spec": ["名称", "规格", "型号", "description", "desc"],
        "qty": ["数量", "quantity", "qty"],
        "unit": ["单位", "unit"],
        "price": ["单价", "price", "unitprice", "rmb"],
        "amount": ["金额", "amount", "合计"],
        "date": ["交期", "到货", "交货", "delivery"],
        "remark": ["备注", "comments", "remark"],
    }
    hits = {field for field, keywords in categories.items() if _has_any(text, keywords)}
    value_like_rows = 0
    for row in rows[1:]:
        row_text = " ".join(str(value or "") for value in row)
        non_empty = sum(1 for value in row if str(value or "").strip())
        has_code = bool(re.search(r"\b[A-Za-z0-9]+(?:[-_][A-Za-z0-9]+){1,}\b|\b[A-Za-z0-9]{8,}\b", row_text))
        money_count = len(re.findall(r"\d[\d,]*(?:\.\d+)?", row_text))
        has_date = bool(re.search(r"20\d{2}[/-]\d{1,2}[/-]\d{1,2}", row_text))
        if non_empty >= 5 and (has_code or has_date) and money_count >= 2:
            value_like_rows += 1
    has_identity = bool({"identity", "spec"}.intersection(hits))
    value_field_count = len({"qty", "unit", "price", "amount", "date"}.intersection(hits))
    score = len(hits)
    if value_like_rows >= 2 and max_cols >= 5:
        score += 2
    return score, has_identity, value_field_count


def _classify_docling_table(table: dict[str, Any]) -> str:
    rows = table.get("rows") or []
    if not rows:
        return "unknown"
    max_cols = max((len(row) for row in rows), default=0)
    text = _docling_compact_text(table)
    detail_score, has_identity, value_field_count = _docling_detail_score(table)
    if len(rows) >= 2 and max_cols >= 3 and detail_score >= 4 and has_identity and value_field_count >= 2:
        return "detail_table"

    order_header_keywords = ["订单号", "采购单号", "采购订单号", "合同编号", "合同号", "客户", "供应商", "买方", "卖方", "vendor", "pono", "p.ono", "contractno", "日期"]
    payment_shipping_keywords = ["付款", "月结", "结算", "账期", "收货", "送货", "地址", "联系人", "电话", "传真", "shipto", "paymentterms", "paymentterms", "deliveryto"]
    terms_notes_keywords = ["备注", "条款", "环保", "质量", "验收", "违约", "争议", "合计金额", "大写金额", "总金额", "rohs", "包装", "责任"]

    if _has_any(text, order_header_keywords) and (max_cols <= 4 or len(rows) <= 6):
        return "order_header"
    if _has_any(text, payment_shipping_keywords):
        return "payment_shipping"
    if _has_any(text, terms_notes_keywords):
        return "terms_notes"
    if _has_any(text, order_header_keywords):
        return "order_header"
    return "unknown"


def _classify_docling_tables(docling_result: dict[str, Any]) -> list[dict[str, Any]]:
    classified: list[dict[str, Any]] = []
    for index, table in enumerate(docling_result.get("tables") or []):
        rows = table.get("rows") or []
        table_type = _classify_docling_table(table)
        classified.append(
            {
                "title": table.get("title") or f"Docling 表 {index + 1}",
                "rows": rows,
                "method": "docling_markdown",
                "page_index": "",
                "table_index": index,
                "table_type": table_type,
            }
        )
    return classified


def _docling_source_tables(classified_tables: list[dict[str, Any]]) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    for table in classified_tables:
        if table.get("table_type") != "detail_table":
            continue
        tables.append({**table, "table_index": len(tables)})
    return tables


def _parse_pdf(file_item: dict[str, str]) -> dict[str, Any]:
    source_file = file_item["original_filename"]
    path = Path(file_item["stored_path"])
    document = _document_base(source_file, "pdf")
    native = parse_pdf_native(path)
    docling_result = parse_pdf_with_docling(path, content_sha256=str(file_item.get("content_sha256") or ""))
    markdown = str(docling_result.get("markdown") or "")
    combined_text = "\n".join(part for part in [native.get("text", ""), markdown] if part)
    template = identify_template(source_file, combined_text)
    classified_docling_tables = _classify_docling_tables(docling_result)
    docling_tables = _docling_source_tables(classified_docling_tables)
    document["parser_mode"] = "template_docling_pdf" if template and docling_tables else "docling_pdf" if docling_tables else "pdf_native_fallback"
    if template:
        document["template_id"] = template.template_id
        document["template_label"] = template.label
    document["page_count"] = native["page_count"]
    document["warnings"].extend(native.get("warnings") or [])
    document["pages"] = native["pages"]
    document["order_number"] = likely_order_number(combined_text)
    document["supplier"] = likely_supplier(combined_text, template)
    document["docling_tables"] = docling_result.get("tables") or []
    document["docling_classified_tables"] = classified_docling_tables
    document["source_tables"] = docling_tables
    document["docling_lines"] = docling_result.get("lines") or []
    native_word_count = sum(len(page.get("words") or []) for page in document["pages"])
    native_cell_count = sum(len(table.get("cells") or []) for page in document["pages"] for table in page.get("tables", []))
    if markdown:
        document["markdown"] = markdown
        if docling_tables:
            document["warnings"].append("Docling Markdown 已生成，并作为 PDF 明细表主解析来源。")
        else:
            document["warnings"].append("Docling Markdown 已生成，但未检测到可靠明细表，已回退到模板/pdfplumber 解析。")
    elif docling_result.get("error"):
        docling_error = str(docling_result.get("error"))
        document["warnings"].append(f"Docling Markdown 不可用：{docling_error}")
    used_rendered_ocr = False
    if not docling_tables and native_word_count < 8 and native_cell_count < 8:
        with tempfile.TemporaryDirectory() as temp_dir:
            image_paths = _pdf_to_images(path, Path(temp_dir))
            ocr_pages = []
            for page_index, image_path in enumerate(image_paths):
                ocr_pages.append(parse_image_tables(image_path, page_index=page_index))
            document["pages"] = ocr_pages
            document["page_count"] = len(ocr_pages)
            document["parser_mode"] = "pdf_rendered_page_ocr"
            document["warnings"].append("PDF 原生解析内容过少，已渲染页面并用 OCR 文字框还原。")
            used_rendered_ocr = True
    if native.get("text_quality", {}).get("has_text") and not native.get("text_quality", {}).get("readable") and not used_rendered_ocr:
        document["warnings"].append("检测到 PDF 文字层疑似乱码，未做整页 OCR；结果保留原生坐标和可读数字字段。")
    return document


def _parse_image(file_item: dict[str, str]) -> dict[str, Any]:
    source_file = file_item["original_filename"]
    path = Path(file_item["stored_path"])
    document = _document_base(source_file, "image")
    template = identify_template(source_file, "")
    document["parser_mode"] = "template_image" if template else "image_table_cell_ocr"
    if template:
        document["template_id"] = template.template_id
        document["template_label"] = template.label
        document["supplier"] = likely_supplier("", template)
    page = parse_image_tables(path, page_index=0)
    document["page_count"] = 1
    document["pages"] = [page]
    document["warnings"].extend(page.get("warnings") or [])
    return document


def _pdf_to_images(pdf_path: Path, output_dir: Path) -> list[Path]:
    pdftoppm = shutil.which("pdftoppm") or shutil.which("pdftoppm.cmd")
    if pdftoppm:
        prefix = output_dir / "page"
        try:
            subprocess.run(
                [pdftoppm, "-png", "-r", "180", str(pdf_path), str(prefix)],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            images = sorted(output_dir.glob("page-*.png"))
            if images:
                return images
        except subprocess.CalledProcessError:
            pass

    try:
        import pypdfium2 as pdfium
    except Exception as exc:
        raise RuntimeError("PDF 页面渲染失败，且 pypdfium2 不可用，无法进入 OCR 兜底。") from exc

    pdf = pdfium.PdfDocument(str(pdf_path))
    image_paths: list[Path] = []
    try:
        for index in range(len(pdf)):
            page = pdf[index]
            bitmap = page.render(scale=2.5)
            image = bitmap.to_pil()
            image_path = output_dir / f"page-{index + 1:03d}.png"
            image.save(image_path)
            image_paths.append(image_path)
    finally:
        pdf.close()
    return image_paths


def _parse_scanned_pdf(file_item: dict[str, str]) -> dict[str, Any]:
    source_file = file_item["original_filename"]
    pdf_path = Path(file_item["stored_path"])
    document = _document_base(source_file, "pdf")
    document["parser_mode"] = "pdf_scan_cell_ocr"
    with tempfile.TemporaryDirectory() as temp_dir:
        image_paths = _pdf_to_images(pdf_path, Path(temp_dir))
        pages = []
        for page_index, image_path in enumerate(image_paths):
            pages.append(parse_image_tables(image_path, page_index=page_index))
        document["pages"] = pages
        document["page_count"] = len(pages)
        for page in pages:
            document["warnings"].extend(page.get("warnings") or [])
    return document


def _parse_file(file_item: dict[str, str]) -> dict[str, Any]:
    suffix = Path(file_item["stored_path"]).suffix.lower()
    if suffix == ".pdf":
        document = _parse_pdf(file_item)
        has_cells = any(table.get("cells") for page in document["pages"] for table in page.get("tables", []))
        if has_cells:
            return document
        return _parse_scanned_pdf(file_item)
    if suffix in IMAGE_EXTENSIONS:
        return _parse_image(file_item)
    raise ValueError(f"不支持的文件类型：{file_item['original_filename']}")


def _legacy_to_purchase_document(document: dict[str, Any], *, fallback_reason: str = "") -> dict[str, Any]:
    business_doc = build_business_document(document)
    raw_tables: list[dict[str, Any]] = []
    for table_index, table in enumerate(business_doc.get("source_tables") or []):
        raw_tables.append(
            {
                "page_index": 0,
                "table_index": table_index,
                "bbox": [],
                "rows": table.get("rows") or [],
                "method": table.get("method") or "legacy_fallback",
                "confidence": "",
            }
        )

    legacy_headers = ["序号", "物料编码", "名称规格", "单位", "数量", "单价", "金额", "到货日期", "备注"]
    if not raw_tables and business_doc.get("detail_rows"):
        rows = [legacy_headers]
        for detail in business_doc.get("detail_rows") or []:
            rows.append([str(detail.get(header, "")) for header in legacy_headers])
        raw_tables.append({"page_index": 0, "table_index": 0, "bbox": [], "rows": rows, "method": "legacy_fallback", "confidence": ""})

    mapped_rows: list[dict[str, Any]] = []
    for detail in business_doc.get("detail_rows") or []:
        standard = {
            "序号": detail.get("序号", ""),
            "物料编码": detail.get("物料编码", ""),
            "物料名称": detail.get("名称规格", ""),
            "说明": "",
            "数量": detail.get("数量", ""),
            "单位": detail.get("单位", ""),
            "含税单价": detail.get("单价", ""),
            "金额": detail.get("金额", ""),
            "交货日期": detail.get("到货日期", ""),
            "备注": detail.get("备注", ""),
        }
        mapped_rows.append(
            {
                "original": {key: value for key, value in standard.items() if value},
                "standard": standard,
                "page_index": detail.get("_page_index", 0),
                "table_index": 0,
                "row_index": len(mapped_rows),
                "raw_text": detail.get("_raw_text", ""),
                "confidence": detail.get("_confidence", ""),
                "method": f"legacy_{detail.get('_method', '')}".strip("_"),
            }
        )

    issues = []
    if fallback_reason:
        issues.append(
            {
                "page_index": "",
                "region": "通用 pipeline",
                "field": "",
                "raw_value": "",
                "clean_value": "",
                "confidence": 0,
                "message": f"新 pipeline 失败，已回退旧流程：{fallback_reason}",
            }
        )

    return {
        "pipeline_version": "purchase_order_v1_legacy_fallback",
        "source_file": document.get("source_file", ""),
        "file_type": document.get("file_type", ""),
        "parser_mode": f"legacy_fallback:{document.get('parser_mode', '')}",
        "template_id": document.get("template_id", ""),
        "template_label": document.get("template_label", ""),
        "page_count": business_doc.get("page_count", document.get("page_count", 0)),
        "started_at": document.get("started_at", ""),
        "header_info": business_doc.get("header_info") or {},
        "pages": [],
        "regions": [],
        "raw_detail_tables": raw_tables,
        "mapped_detail_rows": mapped_rows,
        "sections": {
            "备注": business_doc.get("notes") or [],
            "条款": business_doc.get("terms") or [],
            "付款信息": business_doc.get("payment_info") or [],
            "收货信息": business_doc.get("shipping_info") or [],
            "签核区": [],
        },
        "issues": issues,
        "warnings": document.get("warnings") or [],
    }


def _parse_file_with_purchase_pipeline(file_item: dict[str, str], work_dir: Path) -> dict[str, Any]:
    try:
        return run_purchase_order_pipeline(file_item, work_dir)
    except Exception as exc:
        legacy_document = _parse_file(file_item)
        return normalize_purchase_document(_legacy_to_purchase_document(legacy_document, fallback_reason=str(exc)))


def _parse_repair_and_project(
    file_item: dict[str, str],
    work_dir: Path,
    *,
    ai_config: AiRepairConfig,
    log: Callable[[str], None] | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    document = _parse_file_with_purchase_pipeline(file_item, work_dir)
    ai_started = time.perf_counter()
    document = audit_and_repair_purchase_document(document, config=ai_config, log=log)
    document["ai_config_summary"] = ai_config.safe_metadata()
    add_stage_ms(
        "ai_repair",
        (time.perf_counter() - ai_started) * 1000,
        summary=document.setdefault("performance_summary", {}),
    )
    factory_started = time.perf_counter()
    factory_summary = project_factory_document(document, config=ai_config, log=log)
    add_stage_ms(
        "factory_projection",
        (time.perf_counter() - factory_started) * 1000,
        summary=document.setdefault("performance_summary", {}),
    )
    return document, factory_summary


def _write_purchase_results(
    documents: list[dict[str, Any]],
    job_dir: Path,
    bundle_stem: str,
) -> tuple[Path, dict[str, int], list[Path]]:
    if not documents:
        raise RuntimeError("没有成功解析的采购单，未生成结果文件。")
    result_files: list[Path] = []
    aggregate_stats = {
        "page_count": 0,
        "structured_count": 0,
        "ready_count": 0,
        "review_count": 0,
        "issue_count": 0,
    }
    result_dir = job_dir / "factory_results"
    result_dir.mkdir(parents=True, exist_ok=True)
    for index, document in enumerate(documents, start=1):
        source_stem = safe_result_stem(document.get("source_file"), fallback=f"采购单_{index:03d}")
        result_path = result_dir / f"{index:03d}_{source_stem}_采购单转换结果.xlsx"
        excel_started = time.perf_counter()
        document_stats = write_purchase_order_workbook([document], result_path)
        performance_summary = document.setdefault("performance_summary", {})
        add_stage_ms(
            "excel_write",
            (time.perf_counter() - excel_started) * 1000,
            summary=performance_summary,
        )
        stages = performance_summary.setdefault("stage_ms", {})
        stages["total"] = round(
            sum(
                float(stages.get(name, 0.0) or 0.0)
                for name in ["pipeline_total", "ai_repair", "factory_projection", "excel_write"]
            ),
            3,
        )
        result_files.append(result_path)
        for key in aggregate_stats:
            aggregate_stats[key] += int(document_stats.get(key, 0) or 0)

    if len(result_files) == 1:
        output_path = job_dir / f"{bundle_stem}_采购单转换结果.xlsx"
        shutil.copy2(result_files[0], output_path)
    else:
        output_path = job_dir / f"{bundle_stem}_采购单转换结果.zip"
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for result_path in result_files:
                archive.write(result_path, arcname=result_path.name)
    return output_path, aggregate_stats, result_files


def get_or_create_internal_sales_result(job: dict[str, Any]) -> Path:
    stored_result_path = Path(str(job.get("stored_result_path") or ""))
    if not stored_result_path.exists():
        raise FileNotFoundError("原转换结果不存在，无法定位任务目录。")

    job_dir = stored_result_path.parent
    json_dir = job_dir / "json"
    json_paths = sorted(json_dir.glob("*.json")) if json_dir.exists() else []
    if not json_paths:
        raise FileNotFoundError("该任务没有保留可生成内销模板的解析数据，请重新上传原文件。")

    documents: list[dict[str, Any]] = []
    for json_path in json_paths:
        try:
            document = json.loads(json_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise ValueError(f"解析数据读取失败：{json_path.name}") from exc
        if not isinstance(document, dict):
            raise ValueError(f"解析数据格式无效：{json_path.name}")
        documents.append(document)

    suffix = ".xlsx" if len(documents) == 1 else ".zip"
    output_path = job_dir / f"{job_dir.name}_内销导入模板{suffix}"
    if output_path.exists():
        return output_path

    with tempfile.TemporaryDirectory(prefix="internal_sales_", dir=job_dir) as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        generated_files: list[Path] = []
        for index, document in enumerate(documents, start=1):
            source_stem = safe_result_stem(document.get("source_file"), fallback=f"采购单_{index:03d}")
            workbook_path = temp_dir / f"{index:03d}_{source_stem}_内销导入模板.xlsx"
            write_internal_sales_workbook(document, workbook_path)
            generated_files.append(workbook_path)

        temp_output = temp_dir / output_path.name
        if len(generated_files) == 1:
            shutil.copy2(generated_files[0], temp_output)
        else:
            with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                for workbook_path in generated_files:
                    archive.write(workbook_path, arcname=workbook_path.name)
        temp_output.replace(output_path)
    return output_path


def run_pdf_excel_job(job_id: int, employee_id: str) -> None:
    job = get_job(job_id)
    if not job or job["employee_id"] != employee_id:
        return

    update_job_status(job_id, status="running", log_text="", current_row=0)
    manifest_path = Path(job["stored_input_path"])
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    files = manifest.get("files") or []
    max_workers = max(1, min(int(manifest.get("max_workers") or 2), 4))
    ai_config = config_from_manifest_snapshot(manifest.get("ai_config"))
    job_dir = manifest_path.parent
    json_dir = job_dir / "json"
    work_dir = job_dir / "purchase_pipeline"
    json_dir.mkdir(parents=True, exist_ok=True)
    work_dir.mkdir(parents=True, exist_ok=True)

    append_job_log(job_id, f"开始 PDF/图片转Excel任务，共 {len(files)} 个文件。", total_rows=len(files), current_row=0)
    append_job_log(job_id, ai_config.safe_status())
    documents: list[dict[str, Any]] = []
    success_count = 0
    fail_count = 0
    log_lock = Lock()

    def safe_job_log(message: str, **kwargs: Any) -> None:
        with log_lock:
            append_job_log(job_id, message, **kwargs)

    try:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_map = {
                executor.submit(
                    _parse_repair_and_project,
                    file_item,
                    work_dir / Path(file_item["stored_path"]).stem,
                    ai_config=ai_config,
                    log=lambda message, filename=file_item["original_filename"]: safe_job_log(
                        f"{filename}：{message}"
                    ),
                ): file_item
                for file_item in files
            }
            for future in as_completed(future_map):
                file_item = future_map[future]
                try:
                    document, factory_summary = future.result()
                    if document.get("layout_cache_hit"):
                        safe_job_log(f"{file_item['original_filename']}：已命中历史版式缓存。")
                    documents.append(document)
                    success_count += 1
                    performance = document.get("performance_summary") or {}
                    safe_job_log(
                        f"完成：{file_item['original_filename']}（{document['parser_mode']}，{document['page_count']} 页；"
                        f"厂内可导入 {factory_summary.get('ready_rows', 0)} 行，需复核 {factory_summary.get('review_rows', 0)} 行；"
                        f"解析 {float((performance.get('stage_ms') or {}).get('pipeline_total') or 0) / 1000:.2f}s）",
                        success_count=success_count,
                        fail_count=fail_count,
                        current_row=success_count + fail_count,
                        total_rows=len(files),
                    )
                except Exception as exc:
                    fail_count += 1
                    safe_job_log(
                        f"失败：{file_item['original_filename']}：{exc}",
                        success_count=success_count,
                        fail_count=fail_count,
                        current_row=success_count + fail_count,
                        total_rows=len(files),
                    )

        documents.sort(key=lambda document: document.get("source_file", ""))
        output_path, aggregate_stats, result_files = _write_purchase_results(
            documents,
            job_dir,
            manifest_path.parent.name,
        )
        for index, document in enumerate(documents, start=1):
            source_stem = safe_result_stem(document.get("source_file"), fallback=f"采购单_{index:03d}")
            json_path = json_dir / f"{index:03d}_{source_stem}.json"
            json_path.write_text(json.dumps(document, ensure_ascii=False, indent=2), encoding="utf-8")
        append_job_log(
            job_id,
            f"结果已生成：{output_path.name}，采购单 {len(result_files)} 份、"
            f"客户结果表 {aggregate_stats['page_count']} 页、明细 {aggregate_stats['structured_count']} 行；"
            f"厂内可导入 {aggregate_stats['ready_count']} 行，需复核 {aggregate_stats['review_count']} 行。",
        )
        update_job_status(
            job_id,
            status="completed" if fail_count == 0 else "failed",
            stored_result_path=str(output_path),
            success_count=success_count,
            fail_count=fail_count,
            skip_count=0,
            current_row=success_count + fail_count,
            total_rows=len(files),
            error_message=None if fail_count == 0 else "部分文件转换失败，请查看日志。",
            completed=True,
        )
    except Exception as exc:
        append_job_log(job_id, f"PDF/图片转Excel任务失败：{exc}")
        update_job_status(job_id, status="failed", error_message=str(exc), completed=True)
        raise
