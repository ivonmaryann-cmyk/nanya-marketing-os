from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_pending_risks import match_pending_formal_risk


def main() -> None:
    assert match_pending_formal_risk("惠州智恩", "NY-A2 0.25mm")["risk_id"] == "PFR-001"
    assert match_pending_formal_risk("广东依顿", "NY2150H ANTI-CAF")["risk_id"] == "PFR-004"
    assert match_pending_formal_risk("广东依顿", "NY2150 ANTI-CAF") is None
    assert match_pending_formal_risk("湖奥士康", "NY2150 耐CAF")["risk_id"] == "PFR-005"
    assert match_pending_formal_risk("广州名幸", "NY2140L A级假板")["risk_id"] == "PFR-006"
    assert match_pending_formal_risk("普通客户", "NY-A2") is None
    print("pending formal risks smoke passed")


if __name__ == "__main__":
    main()
