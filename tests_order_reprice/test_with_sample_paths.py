from __future__ import annotations

import os
import sys
from pathlib import Path


PACKAGE_OR_PROJECT_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = PACKAGE_OR_PROJECT_ROOT if (PACKAGE_OR_PROJECT_ROOT / "fangzheng_web_app").exists() else Path.cwd()
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


from fangzheng_web_app.order_reprice_service import (  # noqa: E402
    find_quote_price,
    load_quote_rows,
    process_block1,
    process_block2,
    process_block3,
)


def sample_root() -> Path:
    return Path(os.environ.get("ORDER_REPRICE_SAMPLE_ROOT", r"D:\桌面\胜宏"))


def require_dir(path: Path) -> bool:
    if not path.exists():
        print(f"SKIP: sample directory not found: {path}")
        return False
    return True


def excel_files(path: Path) -> list[Path]:
    return sorted(path.glob("*.xls*"))


def find_file(path: Path, keyword: str) -> Path | None:
    for file_path in excel_files(path):
        if keyword in file_path.name:
            return file_path
    return None


def run_block1(root: Path) -> None:
    path = root / "功能一原始数据"
    if not require_dir(path):
        return
    customer = find_file(path, "客户")
    factory = find_file(path, "厂内")
    if not customer or not factory:
        print(f"SKIP block1: expected 客户/厂内 Excel in {path}")
        return
    result = process_block1(customer, factory)
    print("block1:", {key: result[key] for key in ("total_rows", "success_count", "fail_count")})


def run_block2(root: Path) -> None:
    path = root / "功能二原始数据"
    if not require_dir(path):
        return
    customer = find_file(path, "客户")
    factory = find_file(path, "厂内")
    quotes = [file_path for file_path in excel_files(path) if file_path not in {customer, factory}]
    if not customer or not factory or not quotes:
        print(f"SKIP block2: expected 客户/厂内/报价 Excel in {path}")
        return
    quote_rows = load_quote_rows(quotes)
    quote = find_quote_price("NY2170P:2116RC55%300M/R", quote_rows)
    if round(float(quote.get("price") or 0), 2) != 5691.00:
        raise AssertionError(f"NY2170P expected 5691, got {quote}")
    alias_quote = find_quote_price("NY3170M2P:2116RC55%200M/R", quote_rows)
    if alias_quote.get("price") is None:
        raise AssertionError(f"NY3170M2P alias should hit quote, got {alias_quote}")
    result = process_block2(customer, factory, quotes)
    print("block2:", {key: result[key] for key in ("total_rows", "success_count", "fail_count")})


def run_block3(root: Path) -> None:
    path = root / "功能三原始数据"
    if not require_dir(path):
        return
    customer = find_file(path, "客户")
    factory = find_file(path, "cxmr") or find_file(path, "厂内")
    if not customer or not factory:
        print(f"SKIP block3: expected 客户/cxmr or 厂内 Excel in {path}")
        return
    result = process_block3(customer, factory)
    print("block3:", {key: result[key] for key in ("total_rows", "success_count", "fail_count")})


def main() -> int:
    root = sample_root()
    if not root.exists():
        print(f"SKIP: ORDER_REPRICE_SAMPLE_ROOT not found: {root}")
        return 0
    run_block1(root)
    run_block2(root)
    run_block3(root)
    print("OK: sample path regression finished")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
