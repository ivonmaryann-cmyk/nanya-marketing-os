from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_service import _load_runtime, analyze_spec


def quote(customer_code: str, customer: str, spec: str) -> dict:
    engine, tables, rules, mappings, _, _ = _load_runtime()
    return analyze_spec(
        engine,
        tables,
        rules,
        spec,
        agent_mapping_tables=mappings,
        customer=customer,
        customer_code=customer_code,
        excel_row=2,
    )


def main() -> None:
    bomin = quote(
        "",
        "江苏博敏",
        "南亚新材料 NY3150HF 1.5 mm J/J HTE/HTE 经690*纬615 mm 含铜 无卤 7628x8 TG150 浅黄色 无水印 CTI≥175",
    )
    assert bomin["formal_code"].split("*", 1)[0] == "AH01500JJ27172422RWADT", bomin
    assert bomin["engine_steps"]["customer_order_rule"], bomin

    f7 = quote("103891", "方正F7", "NY2150 1.240mm 1/1 41x49 有卤 HTE 7x7628")
    assert f7["formal_code"].split("*", 1)[0] == "2B013101141004900YWA1T", f7
    assert f7["engine_steps"]["order_mm"] == 1.31, f7
    assert f7["engine_steps"]["step8_tc_code"] == "T", f7

    hanyu_42 = quote("104158", "江苏瀚宇", "NY2150 042 H/H 43*49 6*7628")
    assert hanyu_42["formal_code"].split("*", 1)[0] == "2B01067HH43004900YWA1T", hanyu_42
    hanyu_4p5 = quote("104158", "江苏瀚宇", "NY2150 4P5 2/2 HTE/HTE 43*49 1*2116")
    assert hanyu_4p5["formal_code"].split("*", 1)[0] == "2B001142243004900YWA1C", hanyu_4p5
    hanyu_31 = quote("104158", "江苏瀚宇", "NY2150 031 1/1 41*49 4*7628")
    assert hanyu_31["status"] == "失败" and "无法识别厚度" in hanyu_31["reason"], hanyu_31

    taixing = quote("", "泰兴电路", "CCL FR4.0 0.102 HV1/HV1 0.172 43X49 1J NY6300(C)")
    assert taixing["status"] == "成功", taixing
    assert taixing["engine_steps"]["step2_thick_code"] == "00102", taixing
    assert taixing["engine_steps"]["step8_tc_code"] == "C", taixing

    meizhou = quote("", "梅州奔创", "NY3150HF 0.1 H/H 41*49 A级 不含铜 黄芯 无标 无卤 TG150 2116*1")
    assert meizhou["engine_steps"]["step1_glue_code"] == "AH", meizhou
    assert meizhou["engine_steps"]["step5_glue_cat_code"] == "R", meizhou

    zhongfu = quote("103683", "鹤山中富", "FR4_NY2170H南亚0.23-2/2不含铜_黄无41*49内无")
    assert zhongfu["engine_steps"]["step7_grade_code"] == "AP", zhongfu
    shenzhen_zhongfu = quote("103067", "深圳中富", "FR4_NY2170H南亚0.23-2/2不含铜_黄无41*49内无")
    assert shenzhen_zhongfu["engine_steps"]["step7_grade_code"] == "AP", shenzhen_zhongfu

    jian_shengyi = quote("", "吉安生益", "NY2150 0.8mm 1/1 41*49 HTE AT板")
    assert jian_shengyi["engine_steps"]["step7_grade_code"] == "AC", jian_shengyi

    print("20260721 aggregate deterministic fixes smoke: PASS")


if __name__ == "__main__":
    main()
