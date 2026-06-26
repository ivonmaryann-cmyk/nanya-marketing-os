from __future__ import annotations

import sys
from pathlib import Path


PACKAGE_OR_PROJECT_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = PACKAGE_OR_PROJECT_ROOT if (PACKAGE_OR_PROJECT_ROOT / "fangzheng_web_app").exists() else Path.cwd()
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


from fangzheng_web_app.order_reprice_service import (  # noqa: E402
    _block3_factory_price,
    _is_pp_spec,
    _number_choice_matches,
    _pp_length_from_spec,
    _rc_matches,
    _token_matches,
)


class FakeRow(dict):
    def get(self, key, default=None):
        return super().get(key, default)


def assert_true(value, message: str) -> None:
    if not value:
        raise AssertionError(message)


def assert_equal(left, right, message: str) -> None:
    if left != right:
        raise AssertionError(f"{message}: {left!r} != {right!r}")


def main() -> int:
    assert_true(_is_pp_spec("NY-P3P(C):1078RC66%300M/R"), "NY-P3P(C) should be PP")
    assert_equal(_pp_length_from_spec("NY2170P:2116RC55%300M/R"), 300.0, "PP length M/R")
    assert_equal(_pp_length_from_spec("NY6300SP 1078 含量64.00% 卷300M"), 300.0, "PP length 卷300M")

    assert_true(_rc_matches("64%-66%", 64), "RC lower bound")
    assert_true(_rc_matches("64%-66%", 65), "RC middle")
    assert_true(_rc_matches("64%-66%", 66), "RC upper bound")
    assert_true(not _rc_matches("64%-66%", 67), "RC outside range")
    assert_true(_rc_matches("0.64-0.66", 65), "decimal RC range")

    assert_true(_token_matches("2313/3313", "2313"), "glass option 2313")
    assert_true(_token_matches("2313/3313", "3313"), "glass option 3313")

    assert_true(_number_choice_matches("0.100/0.102", 0.100, 0.003), "thickness option 0.100")
    assert_true(_number_choice_matches("0.100/0.102", 0.102, 0.003), "thickness option 0.102")

    price, note, item = _block3_factory_price(
        FakeRow({"项次": "3", "单价": 106, "规格": "NY-P3P(C) 1078 含量66.00% 卷300M"}),
        "NY-P3P(C):1078RC66%300M/R",
    )
    assert_equal(price, 31800.0, "block3 PP converted price")
    assert_equal(item, "3", "factory item")
    assert_true("PP" in note, "conversion note should mention PP")

    price, note, _item = _block3_factory_price(
        FakeRow({"项次": "1", "单价": 181.72, "规格": "NY6300S 0.102mm 1/1 37x43"}),
        "NY6300S:0.102mm1/1(RTF2)不含铜37*43(1078*1)",
    )
    assert_equal(price, 181.72, "block3 non-PP price")
    assert_true("单价" in note, "non-PP note should mention 单价")

    print("OK: order_reprice smoke tests passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
