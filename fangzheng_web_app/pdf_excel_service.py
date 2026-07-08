from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any

from werkzeug.utils import secure_filename

from .db import append_job_log, create_job, get_job, update_job_status
from .docling_parser import parse_pdf_with_docling
from .image_table_parser import parse_image_tables
from .job_control import launch_job_process
from .json_to_excel import build_business_document, write_conversion_workbook
from .paths import JOBS_DIR
from .pdf_native_parser import parse_pdf_native
from .purchase_excel_writer import write_purchase_order_workbook
from .purchase_order_pipeline import run_purchase_order_pipeline
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
        manifest_files.append(
            {
                "original_filename": original_filename,
                "stored_path": str(stored_path),
                "extension": suffix,
            }
        )

    manifest_path = job_dir / "manifest.json"
    manifest = {
        "feature": FEATURE,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "max_workers": max(1, min(int(max_workers or 2), 4)),
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
    docling_result = parse_pdf_with_docling(path)
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
        return _legacy_to_purchase_document(legacy_document, fallback_reason=str(exc))


def run_pdf_excel_job(job_id: int, employee_id: str) -> None:
    job = get_job(job_id)
    if not job or job["employee_id"] != employee_id:
        return

    update_job_status(job_id, status="running", log_text="", current_row=0)
    manifest_path = Path(job["stored_input_path"])
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    files = manifest.get("files") or []
    max_workers = max(1, min(int(manifest.get("max_workers") or 2), 4))
    job_dir = manifest_path.parent
    json_dir = job_dir / "json"
    work_dir = job_dir / "purchase_pipeline"
    json_dir.mkdir(parents=True, exist_ok=True)
    work_dir.mkdir(parents=True, exist_ok=True)

    append_job_log(job_id, f"开始 PDF/图片转Excel任务，共 {len(files)} 个文件。", total_rows=len(files), current_row=0)
    documents: list[dict[str, Any]] = []
    success_count = 0
    fail_count = 0

    try:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_map = {
                executor.submit(
                    _parse_file_with_purchase_pipeline,
                    file_item,
                    work_dir / Path(file_item["stored_path"]).stem,
                ): file_item
                for file_item in files
            }
            for future in as_completed(future_map):
                file_item = future_map[future]
                try:
                    document = future.result()
                    documents.append(document)
                    success_count += 1
                    json_path = json_dir / f"{Path(file_item['stored_path']).stem}.json"
                    json_path.write_text(json.dumps(document, ensure_ascii=False, indent=2), encoding="utf-8")
                    append_job_log(
                        job_id,
                        f"完成：{file_item['original_filename']}（{document['parser_mode']}，{document['page_count']} 页）",
                        success_count=success_count,
                        fail_count=fail_count,
                        current_row=success_count + fail_count,
                        total_rows=len(files),
                    )
                except Exception as exc:
                    fail_count += 1
                    append_job_log(
                        job_id,
                        f"失败：{file_item['original_filename']}：{exc}",
                        success_count=success_count,
                        fail_count=fail_count,
                        current_row=success_count + fail_count,
                        total_rows=len(files),
                    )

        documents.sort(key=lambda document: document.get("source_file", ""))
        output_path = job_dir / f"{manifest_path.parent.name}_PDF图片转Excel结果.xlsx"
        output_path = job_dir / f"{manifest_path.parent.name}_采购单转换结果.xlsx"
        stats = write_purchase_order_workbook(documents, output_path)
        append_job_log(
            job_id,
            f"Excel 已生成：{output_path.name}，客户结果表 {stats.get('page_count', 0)} 页，明细 {stats.get('structured_count', 0)} 行。",
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
