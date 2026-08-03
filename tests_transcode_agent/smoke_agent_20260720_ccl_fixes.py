from pathlib import Path
import sys

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fangzheng_web_app.transcode_agent_service import _customer_matches, _load_runtime, analyze_spec


CASES = [
    ("103295", "惠州智恩", "NY-A2 0.8mm H/H 41*49 HTE", "AL00800HH41004900RWACT"),
    ("123015", "赣州超跃", "NY3170HF 0.8mm 1/1 41*49 HTE", "BD008001141004900RWACT"),
    ("", "深圳联创", "NY3150HC 0.8mm 1/1 41*49 HTE", "3H008001141004900YWACT"),
    ("103901", "广东依顿", "覆铜板10±1MIL H/HOZ 74*49 TG150 HTE 不连铜 NY2150H 2张2116 ANTI-CAF", "2H00254HH74004900RWA1C"),
    ("106008", "川英创力", "NY2150 0.8mm 1/1 41*49 TG150 HTE", "2B008001141004900YWA1T"),
    ("123015", "赣州超跃", "NY2140 0.8mm H/H 41*49 HTE", "2A00800HH41004900YWF1T"),
    ("123015", "赣州超跃", "NY2140 0.8mm 1/1 41*49 HTE", "2A008001141004900YWA1T"),
    ("103898", "广东喜珍", "NY2150 0.8mm 1/1 41*49 耐CAF HTE", "2B008001141004900YWA1T"),
    ("", "广州名幸", "覆铜板 NY2140L 0.80 H/H 41*49 A级假板", "2K00800HH41004900RWA1T"),
    ("232005", "台湾敬鹏", "NYHP-7350 0.0200\" H/H 43”x49\"", "7L00508HH43004900RRA1C"),
    ("103507", "众信天成", "南亚 NY2150 0.8 2/2 41*49 TG150", "2B008002241004900YWA1T"),
    ("103507", "众信天成", "南亚 NY2150 0.8 2/2 37*49 TG150", "2B008002237004900YWA1T"),
    ("103786", "惠州泰和", "NY3170HF 5mil 不含铜 5/5 43\"x49\" HTE 无卤", "BD001275543004900RWA1C"),
    ("193012", "鹤山超前", "NY6300S 0.100mm 芯厚 1/1 RTF2+RTF2 C 无水印 37.00x49.00 HW 0.1*940*1245MM", "6C001001137004900RBATC"),
    ("103894", "珠海崇达", "覆铜板 FR4 NY2150 1.465 H/H HTE/HTE 1.5 86 49 7628*8 黄色 无水印 大料 有卤", "2B01465HH86004900YWA1T"),
    ("257115", "越南名幸", "CCL:0.15t H/H 517*618MM NY3176HF", "AT00150HH20362434RWADC"),
    ("103831", "东莞福哥", "NY3150HFP 1.0mm 0/0 41*49 黄色 无水印 A级 总厚 无卤素 铝膜生产", "3B010000041004900YWA6T"),
    ("203011", "莞森玛仕", "HW 0.076MM(1078*1) H/H NY6300 21.3*24.3\"", "6W00076HH21302430ROA1C"),
    ("193055", "深三德冠", "NY3170M 16mil T/T  7628* 37\"*49\" (帮裁切:412mm纬向*468mm经向) 南亚", "3M00406TT18431623RWA1C"),
    ("257112", "泰国方正", "NY6300 1.000mm H/H 41\"x49\" Halogenated HTE 9x2116 开厂测试用", "6W01030HH41004900RWA1T"),
    ("257112", "泰国方正", "NY3170M2 1.000mm H/H 41\"x49\" Halogenated HTE 9x2116 开厂测试用", "3N01030HH41004900RWA1T"),
    ("104312", "无锡健鼎", "NY3150HF 26mil H/H 41*49(CAF+7628*2+2116*2)无卤素 Non-Dicy Tg150", "AH00660HH41004900RWA1C"),
    ("124016", "龙岩金时", "FR-4 TG170 黄芯 NY3170M2 0.184 1/1 ±0.013 41.2*49.2  (不含铜 0.114 2116*1)", "3N001141141204920RBA1C"),
    ("103312", "深万基隆", "A4级 芯板 NY2140 1.4 0.5/0.5oz 1.4±0.075 49X86 无LOGO", "2A01430JJ86004900YWA4T"),
    ("103312", "深万基隆", "A4级 芯板 NY2140 1.4 0.5/0.5oz 1.4±0.075 49X82 无LOGO", "2A01430JJ82004900YWA4T"),
    ("103312", "深万基隆", "A4级 芯板 NY2140 1.4 0.5/0.5oz 1.4±0.075 49X74 无LOGO", "2A01430JJ74004900YWA4T"),
    ("133021", "乐凯特科", "覆铜板 0.3(不含铜)FR-4 H/H 41*49 无卤素 TG150 上海南亚 NY3150HF", "AH00300HH41004900RWA1C"),
    ("133021", "乐凯特科", "覆铜板 1.2(含铜)FR-4 1/1 41*49 无卤素 TG150南亚新材/无水印/车载板 NY3150HF", "AH012001141004900RWACT"),
    ("133021", "乐凯特科", "覆铜板 0.5(含铜)FR-4 J/J 41*49 无卤素 TG150上海南亚 NY3150HF", "AH00470JJ41004900RWA1C"),
    ("133021", "乐凯特科", "覆铜板 0.7(含铜)FR-4 J/J 41*49 无卤素 TG150 南亚新材 NY3150HF", "AH00670JJ41004900RWA1C"),
    ("105011", "湖南鹰飞", "0.18-0.22 T/T 无卤板料（含铜箔厚度）41*49 NY3170HF", "BD00180TT41004900RWACC"),
    ("104462", "南通展华", "基板：NY6300S 基板 16mil T/T RTF2 TG200 （1078*4)42.8*48.8\"", "6C00406TT43004900RBPGC"),
    ("107010", "大连阿尔卑斯", "NY1140 0.4mm 0/0 19.69x39.65 自然色 无水印 A级 总厚", "1A004000019693965NWA1T"),
    ("106030", "川睿杰鑫", "南亚NY6300S 0.406 H/H 43*49 TG170 无卤 不含铜 无水印", "6C00406HH43004900RBA1C"),
    ("105011", "湖南鹰飞", "NY8320NS 0.34±0.02 T/T  无卤BT 黑芯板料 (含铜箔厚度）41*49", "8B00310TT41004900BWA1C"),
    ("105011", "湖南鹰飞", "NY8320NS 0.245±0.025 T/T 无卤板料 （含铜箔厚度）41*49", "8B00215TT41004900BWA1C"),
    ("133038", "安徽万奔", "南亚 NY1600 FR-4 1.1 J/J 43*49 含铜 135 无水印 600V", "1L01100JJ43304930YWA2T"),
    ("133038", "安徽万奔", "南亚 NY1600 FR-4 0.9 J/J 41\"*49\" 含铜 135 无水印 600V", "1L00900JJ41304930YWF1T"),
    ("133038", "安徽万奔", "南亚 NY1600 FR-4 1.1 J/J 41.3*49.3 含铜 135 无水印 600V", "1L01100JJ41304930YWA2T"),
    ("133038", "安徽万奔", "南亚 NY1600 FR-4 1.5 J/J 41\"*49\" 含铜 135 无水印 600V", "1L01500JJ41304930YWF1T"),
    ("103993", "东莞长谐", "NY 高速板 0.25mm 1/1 (不含铜) (无水印) TG≥170 (NY6300S) 37*49*黄芯", "6C002501137004900RBA1C"),
    ("193030", "珠志博信", "南亚NY3150HF 不含铜 0.4mm 7628*2 H/H ED 43.3*49.3 TG150 无卤", "AH00400HH43304930RWADC"),
    ("193055", "深三德冠", "NY3170M TG190 0.254mm HF H/H 不含铜 HTE铜 2116*2 尺寸41×49 (帮裁切412mm纬向 508mm径向) 南亚", "3M00254HH41004900RWA1C"),
    ("123099", "江志博信", "基材 1.0 H/H 37*49 NY3150HF（不含铜）", "AH01030HH37004900RWA1T"),
    ("193030", "珠志博信", "南亚NY3150HF 不含铜 0.460 7628*2+1080*1 H/H E 43.3*49.3 TG150 无卤", "AH00460HH43304930RWADC"),
    ("103317", "惠威尔高", "86.3*49.3 南亚 NY3150HF 0.5mm 1/1 中TG(TG150) CTI250-399 不含铜 无水印 86.3*49.3 A级 HTE 7628*2张+2116*1张 无卤素 FR-4 耐CAF", "AH005001186304930RWA1C"),
    ("103317", "惠威尔高", "41*49.3 南亚 NY3150HF 0.5mm 1/1 TG150 CTI≥175 不含铜 无水印 41*49.3 A级 HTE 7628*2张+2116*1张 无卤素 FR-4 耐CAF", "AH005001141004930RWA1C"),
    ("193012", "鹤山超前", "NY6300S 0.254mm 芯厚 H/H RTF2+RTF2 A 无水印 37.30x49.30", "6C00254HH37304930RBA1C"),
    ("133021", "乐凯特科", "覆铜板 0.5(含铜)FR-4 J/J 86*49 无卤素 TG150 南亚新材 NY3150HF", "AH00470JJ86004900RWA1C"),
    ("104312", "无锡健鼎", "NY3150HF 15mil H/H 43*49(TFT+CAF+7628*2) 无卤素 Non-Dicy Tg150", "3B00380HH43004900YWATC"),
    ("257115", "越南名幸", "CCL: 0.4t H/H 41\"*49\"(-0/+2\") NY3176HF", "AT00400HH41004900RWA1C"),
    ("104462", "南通展华", "基板： NY6300S 16mil H/Hoz 42.8\"*48.8\" RTF2", "6C00406HH43004900RBA1C"),
    ("193030", "珠志博信", "南亚NY3150HF 不含铜 0.4mm 7628*2 H/H ED 41.3*49.3 TG150 无卤", "AH00400HH41304930RWA1C"),
    ("123099", "江志博信", "基材 0.064 H/H 43*49/NY3150HF（不含铜）RTF", "AH00064HH43004900RRA1C"),
    ("103317", "惠威尔高", "74*49 南亚 NY3150HF 0.6mm 2/2 中TG(TG150) CTI≥175 含铜 无水印 74*49 A级 HTE 7628*2张+1080*1张 无卤素 FR-4 耐CAF", "AH004702274004900RWA1C"),
    ("103067", "深圳中富", "FR4_NY2150H南亚1.10-2/2含铜_黄_无_37.3*49.3_HTE_7628*5", "2W011002237304930RWAPT"),
    ("103067", "深圳中富", "FR4_NY2170H南亚1.00-1/1含铜_黄无37*49内无", "2E010001137004900RWAPT"),
    ("103683", "鹤山中富", "FR4_NY2150H南亚0.36-2/2不含铜_黄无41*49内无", "2W003602241004900RWAPC"),
    ("104354", "常熟敬鹏", "FML NY2170 (FR-4) 无铅 CAF TG170 0.71(0.028\") 2/2 82\"X49\"(2090*1245MM)", "2C007102282304900YWACC"),
    ("", "深南电路", "覆铜板 FR4.0 0.089 HV2H/HV2H 0.124 36X48 2A NY-P1", "6Y00089HH37004900RPA1C"),
    ("", "深南电路", "覆铜板 FR4.1 0.800 R3H/R3H 0.835 36X48 8J NY6666SE", "C800835HH37004900RAA1T"),
]


def main() -> None:
    engine, tables, rules, mappings, _, _ = _load_runtime()

    assert engine.extract_customer_dual_thickness("深南电路", "覆铜板 FR4.0 0.089 HV2H/HV2H 0.124 36X48")[0] == 0.089
    assert engine.extract_customer_dual_thickness("深南电路", "覆铜板 FR4.1 0.800 R3H/R3H 0.835 36X48")[0] == 0.835
    assert engine.extract_customer_dual_thickness("泰兴电路", "CCL FR4.0 0.102 HV1/HV1 0.172 43X49")[0] == 0.102
    assert engine.extract_thickness_mm('沪士 NY-P4 0040" H/H 459X612')[0] == 0.102
    assert engine.extract_thickness_mm('RF-4---0.15 H/H 41*49')[0] == 0.15
    assert engine.extract_thickness_mm('NY3170M 0.152*940*1245MM')[0] == 0.152
    assert engine.extract_copper_spec("覆铜板 FR4.0 0.089 HV2H/HV2H 0.124") == "H/H"
    assert engine.extract_copper_spec("覆铜板 FR4.1 0.127 R21/S2 0.232") == "1/2"
    assert engine.extract_size("CCL 0.114mm 1/H 43''x49''") == (43.0, 49.0)
    assert engine.extract_size("经41*纬49 inch") == (41.0, 49.0)
    assert engine.extract_size("82IN 49IN") == (82.0, 49.0)

    no_code_template = pd.DataFrame([
        ["客户简称", "品名", "客户规格"],
        ["川华兴宇", "2B008001141004900YWA1TA", "NY2150 0.8mm 1/1 41*49"],
    ])
    assert engine.detect_customer_code_column(no_code_template) is None
    coded_template = pd.DataFrame([
        ["客户编号", "客户简称", "客户规格"],
        ["106011", "川华兴宇", "NY2150 0.8mm 1/1 41*49"],
    ])
    assert engine.detect_customer_code_column(coded_template) == 0

    coded_rule = {"客户代码": "106011", "客户简称": "川华兴宇"}
    assert _customer_matches(coded_rule, "106011", "川华兴宇")
    assert _customer_matches(coded_rule, "", "川华兴宇")
    assert _customer_matches(coded_rule, "2B008001141004900YWA1TA", "川华兴宇")
    assert not _customer_matches(coded_rule, "106012", "川华兴宇")
    assert not _customer_matches(coded_rule, "", "川华兴")

    normalized_fallback = analyze_spec(
        engine,
        tables,
        rules,
        "南亚 NY6300S 不含铜 H/HOZ 无卤 0.254*940*1245MM RTF2",
        agent_mapping_tables=mappings,
        customer="鹤山超前",
        customer_code="193012",
        parse_fallback_text="NY6300S 0.254mm 芯厚 H/H RTF2+RTF2 A 无水印 37.30x49.30",
    )
    assert (normalized_fallback.get("candidate_code") or "").split("*")[0][:22] == "6C00254HH37304930RBA1C"
    assert normalized_fallback.get("engine_steps", {}).get("context_fallback_used") is True

    failures = []
    for customer_code, customer, spec, expected in CASES:
        analysis = analyze_spec(
            engine,
            tables,
            rules,
            spec,
            agent_mapping_tables=mappings,
            customer=customer,
            customer_code=customer_code,
        )
        actual = (analysis.get("candidate_code") or "").split("*")[0][:22]
        if actual != expected:
            failures.append((customer, expected, actual, spec))
    assert not failures, "\n".join(str(item) for item in failures)
    print(f"20260720 CCL deterministic fixes smoke: PASS ({len(CASES)} cases)")


if __name__ == "__main__":
    main()
