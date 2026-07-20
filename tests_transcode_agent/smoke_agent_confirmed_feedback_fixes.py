from pathlib import Path
import sys
from tempfile import TemporaryDirectory

from werkzeug.datastructures import FileStorage


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app import db
from fangzheng_web_app import transcode_agent_rules as agent_rules
from fangzheng_web_app import transcode_rules
from fangzheng_web_app.transcode_agent_service import calculate_transcode_agent_quote


DRAFT_PATH = ROOT / "docs/develop0707/客户特殊规则结构化草稿_按原表行_20260708.xlsx"


def main() -> None:
    with TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        original_db_path = db.DATABASE_PATH
        original_agent_rules_dir = agent_rules.TRANSCODE_AGENT_RULES_DIR
        original_agent_versions_dir = agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR
        original_base_versions_dir = transcode_rules.TRANSCODE_RULES_VERSIONS_DIR
        try:
            db.DATABASE_PATH = temp_dir / "storage/app.db"
            db.DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
            db.init_db()

            transcode_rules.TRANSCODE_RULES_VERSIONS_DIR = temp_dir / "rules/transcode/versions"
            transcode_rules.TRANSCODE_RULES_VERSIONS_DIR.mkdir(parents=True, exist_ok=True)
            transcode_rules.ensure_default_transcode_rule_version()

            agent_rules.TRANSCODE_AGENT_RULES_DIR = temp_dir / "rules/transcode_agent"
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR = agent_rules.TRANSCODE_AGENT_RULES_DIR / "versions"
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR.mkdir(parents=True, exist_ok=True)

            with DRAFT_PATH.open("rb") as source:
                upload = FileStorage(
                    stream=source,
                    filename=DRAFT_PATH.name,
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                agent_rules.save_new_transcode_agent_rule_version(
                    upload,
                    updated_by="smoke",
                    remark="confirmed feedback fixes smoke",
                )

            _assert_grade(
                "广东依顿",
                "NY2150 1.45mm H/H 86*49 HTE",
                "A1",
            )
            _assert_grade(
                "广东依顿",
                "NY2150 汽车板 1.45mm H/H 86*49 HTE",
                "AC",
            )
            _assert_field(
                "广东依顿",
                "NY2150H 8mil 1/1 86*49 HTE 芯厚",
                "厚度",
                "00203",
            )
            _assert_field(
                "森德科技",
                "NY2150 0.71mm 2/2 86.3*49.3 HTE",
                "铜厚",
                "22",
                customer_code="257103",
            )
            _assert_field(
                "森德科技",
                "NY2150 0.71mm R/R 86.3*49.3 HTE",
                "铜厚",
                "FF",
                customer_code="257103",
            )
            quote = calculate_transcode_agent_quote(
                "NY2150 1.200mm 总厚 1/1 HTE+HTE A 无水印 41.10*49.10",
                customer="珠海乐健",
                customer_code="103576",
            )
            assert _field_code(quote, "尺寸") == "41104910", quote
            assert _field_code(quote, "基板级别") == "AC", quote
            _assert_grade(
                "珠海乐健",
                "NY2150 1.1mm H/H 37*49 HTE",
                "A1",
                customer_code="103576",
            )
            _assert_field(
                "惠州特创",
                "NY2150 0.100mm 芯厚 1/1 HTE+HTE A 无水印 37.30*49.30",
                "尺寸",
                "37304930",
                customer_code="103613",
            )
            _assert_field(
                "生益电子",
                "NY2150 0.8mm H/H 37*49 HTE",
                "胶系",
                "AS",
                customer_code="103738",
            )
            _assert_field(
                "生益电子",
                "NY2150H 0.8mm H/H 37*49 HTE",
                "胶系",
                "2H",
                customer_code="103738",
            )
            _assert_field(
                "深万基隆",
                "NY2140 0.67mm H/H 43*49 HTE",
                "铜厚",
                "JJ",
                customer_code="103312",
            )
            _assert_field(
                "广东依顿",
                "NY2150 1.500mm 总厚 1.5/1.5 HTE+HTE A 无水印 86.00*49.00",
                "铜厚",
                "FF",
            )
            _assert_field(
                "广东依顿",
                "NY2150 0.8mm H/1 37*49 HTE",
                "铜厚",
                "H1",
            )
            _assert_field(
                "台湾敬鹏",
                "NY6300S 0.510mm 芯厚 H/H HVLP1+HVLP1 A 无水印 43.00*49.00",
                "铜箔类型",
                "O",
            )
            _assert_field(
                "湖北健鼎",
                "NY3150HF 3mil H/H 43.30*49.30 HTE",
                "厚度",
                "00075",
                customer_code="122013",
            )
            _assert_field(
                "无锡健鼎",
                "NY3150HF 2.5mil H/H 41.00*49.00 HTE",
                "厚度",
                "00060",
                customer_code="104312",
            )
            _assert_field(
                "无锡健鼎",
                "NY3150HF 4mil H/H 37.30*49.30 HTE",
                "厚度",
                "00100",
                customer_code="104312",
            )
            _assert_grade(
                "中宝悦嘉",
                "NY1600 0.06mm 1/1 49.0*43.0 HTE",
                "A2",
                customer_code="103006",
            )
            _assert_field(
                "广华升鑫",
                "NY-A2 0.47mm 2/2 41*49 HTE",
                "胶系",
                "AL",
                customer_code="103990",
            )
            _assert_grade(
                "广华升鑫",
                "NY-A2 0.47mm 2/2 41*49 HTE",
                "AC",
                customer_code="103990",
            )
            _assert_field(
                "深圳安比",
                "Lastra FR4 Std. 0.70mm 35/35 um NY2140 1092x1245cm",
                "厚度",
                "00640",
                customer_code="203012",
            )
            _assert_field(
                "深圳安比",
                "NY2140 Lastra FR4 Std. 1.00mm 70/70um 1092x1245mm",
                "胶系",
                "2A",
                customer_code="203012",
            )
            _assert_field(
                "惠州特创",
                "FR4 NY2150汽车专用 0.100 1/1 TG150 940 1245 不含铜 无水印 1*2116 有卤 黄色 CTI值:175",
                "尺寸",
                "37304930",
                customer_code="103613",
            )
            _assert_field(
                "淮安特创",
                "FR4 NY2150汽车专用 1.100 1/1 TG150 1880 1245 含铜 无水印 6*7628 有卤 黄色 CTI值:175",
                "尺寸",
                "74304930",
                customer_code="104359",
            )
            _assert_field(
                "湖奥士康",
                "有卤基板|NY-A1|0.200|1/1|86.3×49.3|不含铜|TG150|HTE|耐CAF|1|上海南亚|7628×1||FR4|黄料|无水印||||",
                "胶系",
                "RC",
                customer_code="105007",
            )
            _assert_grade(
                "江西景旺",
                "CCL NY-A1 0.23mm 1/1 (不含铜) 74inX49in 2116*2 HTE",
                "AC",
                customer_code="123018",
            )
            _assert_field(
                "江西景旺",
                "CCL NY-A1 0.23mm 1/1 (不含铜) 74inX49in 2116*2 HTE",
                "胶系",
                "RC",
                customer_code="123018",
            )
            _assert_field(
                "深万基隆",
                "A级 芯板 FR-4 NY2140 1.4 0.5/0.5Oz 1.4±0.075 49.3X86.3 无LOGO",
                "铜厚",
                "JJ",
                customer_code="103312",
            )
            _assert_field(
                "深万基隆",
                "A级 芯板 FR-4 NY2140 1.4 0.5/0.5Oz 1.4±0.075 49.3X86.3 无LOGO",
                "厚度",
                "01430",
                customer_code="103312",
            )
            _assert_grade(
                "深万基隆",
                "A级 芯板 FR-4 NY2140 1.4 0.5/0.5Oz 1.4±0.075 49.3X86.3 无LOGO",
                "A1",
                customer_code="103312",
            )
            _assert_grade(
                "广东依顿",
                "覆铜板 8±0.8MIL 1/1OZ 86*49\" TG≥150 HTE 不连铜 NY2150H ANTI-CAF",
                "AC",
                customer_code="103901",
            )
            _assert_field(
                "深圳普林",
                "FR4板材 NY2170 1.9mm 1.5/1.5oz 41*49in 含铜 无水印",
                "铜厚",
                "FF",
                customer_code="193047",
            )
            quote_taiwan = calculate_transcode_agent_quote(
                "NY6300S 0.0200\" H/H VLP1 43\"x49\"",
                customer="台湾敬鹏",
                customer_code="232005",
            )
            assert _field_code(quote_taiwan, "厚度") == "00510", quote_taiwan
            assert _field_code(quote_taiwan, "铜箔类型") == "O", quote_taiwan
            _assert_field(
                "黄石广合",
                '敷铜基板 上海南亚 NY3170M2 0.028" C/M级 不含铜 2116*7 1/1 (RTF2/RTF2 /)*经向82"*纬向 49" TG180 无卤素 _',
                "厚度",
                "00711",
                customer_code="122021",
            )
            _assert_field(
                "黄石广合",
                '敷铜基板 上海南亚 NY3170M 0.003" C/M级 不含铜 1078*1 H/H (RTF/RTF/) 经向41"*纬向49" TG180 无卤素_',
                "厚度",
                "00076",
                customer_code="122021",
            )
            quote_chongda = calculate_transcode_agent_quote(
                "覆铜板 FR4 NY3150HC 1.39 3/3 HTE/HTE 1.6 86 49 7628*8 黄色 无水印 大料 无卤",
                customer="崇达一厂",
                customer_code="103769",
            )
            assert _field_code(quote_chongda, "厚度") == "01600", quote_chongda
            assert _field_code(quote_chongda, "基板级别") == "AC", quote_chongda
            _assert_field(
                "兴森快捷",
                "HTg 有卤素 无铅 0.8mm T/C 0.5/0.5OZ HTE 37inch*49(纬)inch (7628*4)",
                "厚度",
                "00830",
                customer_code="103787",
            )
            _assert_field(
                "惠州汇通",
                "NY3150HF 0.5mm 1/1 41x49 黄色 HTE铜箔+无水印 A级 总厚",
                "厚度",
                "00440",
                customer_code="103639",
            )
            _assert_field(
                "广东依顿",
                '覆铜板3±0.5mil 1/HOZ 37*49" TG150 HTE 不连铜 NY2150 Anti-CAF',
                "铜厚",
                "H1",
                customer_code="103901",
            )
            _assert_field(
                "吉安生益",
                '覆铜板 FR-4 NY6300SN 1/2 RTF3/RTF 0.076±0.013mm不含铜 37.00"X49.00"1X1078',
                "铜箔类型",
                "T",
                customer_code="123043",
            )

            print("confirmed feedback fixes smoke passed")
        finally:
            db.DATABASE_PATH = original_db_path
            agent_rules.TRANSCODE_AGENT_RULES_DIR = original_agent_rules_dir
            agent_rules.TRANSCODE_AGENT_RULES_VERSIONS_DIR = original_agent_versions_dir
            transcode_rules.TRANSCODE_RULES_VERSIONS_DIR = original_base_versions_dir


def _assert_grade(customer: str, spec: str, expected_code: str, customer_code: str = "") -> None:
    _assert_field(customer, spec, "基板级别", expected_code, customer_code=customer_code)


def _assert_field(customer: str, spec: str, field: str, expected_code: str, customer_code: str = "") -> None:
    quote = calculate_transcode_agent_quote(spec, customer=customer, customer_code=customer_code)
    actual = _field_code(quote, field)
    assert actual == expected_code, (customer, spec, field, expected_code, quote)


def _field_code(quote: dict, field: str) -> str:
    return next(item["code"] for item in quote["field_evidence"] if item["field"] == field)


if __name__ == "__main__":
    main()
