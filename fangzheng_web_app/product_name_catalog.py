"""New-name standard only; no legacy or customer-specific conversion rules."""
from __future__ import annotations

# category: label, width, product scope. Glue codes are shared by both products.
CATEGORIES = {
    "glue": ("胶系主表", 4, "shared"),
    "copper_type": ("铜箔类型", 1, "board"),
    "copper_weight": ("铜箔规格（oz）", 1, "board"),
    "copper_pair": ("标准双面铜箔组合", 4, "board"),
    "board_spec": ("基板品规", 2, "board"),
    "structure": ("组合结构", 1, "board"),
    "glass_style": ("玻布型号", 4, "pp"),
    "glass_type": ("玻布类型", 1, "pp"),
    "pp_spec": ("PP品规", 2, "pp"),
    "roll_size": ("PP卷料尺寸", 8, "pp"),
    "marking": ("印字", 1, "shared"),
    "board_grade": ("基板品级", 2, "board"),
    "pp_grade": ("PP品级", 2, "pp"),
}
STRUCTURES = {
    "board": [("胶系",4),("厚度",4),("品规",2),("铜箔",4),("组合结构",1),("印字",1),("尺寸",8),("品级",2)],
    "pp": [("胶系",4),("玻布型号",4),("玻布类型",1),("厚度 / RC",3),("品规",2),("印字",1),("尺寸",8),("品级",2)],
}
SEEDS = {
    "copper_pair": {"HHNN":"文档示例：双面H铜重HTE（H单独重量定义待补）","H1RN":"文档示例：非对称铜箔组合","00AA":"铝箔光板","0000":"离型膜","UUNW":"载体薄铜/厚铜接触组合"},
    "copper_type": {"N":"HTE","R":"RTF · Rz≤3μm","F":"RTF2 · Rz≤2.3μm","G":"RTF3 · Rz≤2.1μm","M":"RTF4 · Rz<1.9μm","T":"RTF5 · Rz≤1.5μm","U":"RTF6 · Rz≤1.3μm","V":"HVLP · Rz≤2μm","L":"VLP","B":"HVLP1","P":"HVLP2 · Rz≤1.1μm","Q":"HVLP3 · Rz≤0.9μm","J":"HVLP4 · Rz≤0.7μm","K":"HVLP5 · Rz≤0.5μm","Z":"载体铜箔","C":"RCC","A":"离型膜预留"},
    "copper_weight": {"6":"6","5":"5","4":"4","D":"3.5","A":"3.2","3":"3","G":"2.5","B":"2.2 (75μm)","2":"2","S":"1.8 (63μm)","F":"1.5","L":"1.2 (42μm)","V":"1.1 (38μm)","1":"1","K":"4/5 (28μm)","M":"3/4 (26μm)","N":"0.7 (25μm)","I":"0.64 (22μm)"},
    "board_spec": {"C3":"三级","T3":"三级","C2":"二级","T2":"二级","ZA":"18Ex","ZB":"18Ex(C)","ZC":"18FL","ZE":"18FL(C)","ZF":"T5","ZG":"18GN","ZD":"18Ex-M","ZH":"18FL-M"},
    "structure": {"A":"第一种结构","B":"第二种结构","C":"第三种结构","D":"第四种结构","E":"第五种结构"},
    "glass_style": {x:x for x in ("7638","7667","7630","7628","7625","1506","2118","2116","3313","2313","2113","2112","1086")},
    "glass_type": {"C":"开纤","A":"汽车板玻布","F":"仿布","M":"毛毡","K":"Low Dk PPO","P":"PPO E布","D":"Low Dk2","Q":"石英","T":"Low CTE","H":"Low Dk环氧","Y":"牵引"},
    "pp_spec": {"XX":"通用","GA":"GT+5s","GB":"GT+10s","GX":"自用单张","GY":"单张加严","K5":"深南N05","K7":"深南N07","XY":"多张加严"},
    "roll_size": {"R0004970":"自用49.7英寸","R0004380":"自用43.8英寸","R000XXXX":"外售常规49.5英寸","R000XXXT":"外售43.0英寸","R000XXXE":"外售43.5英寸","R000XXXA":"外售43.3英寸","R0002441":"外售24.41英寸","R000XXXY":"49.5英寸羽边"},
    "marking": {"W":"无水印","Q":"印字","I":"南非印字"},
    "board_grade": {"A0":"A级","B3":"B级","KK":"可靠性","KD":"电性","HW":"688/TFT","AS":"重复品名","PG":"测试","A1":"特殊"},
    "pp_grade": {"A0":"A级","B3":"B级","KK":"可靠性","KD":"电性","HW":"688/TFT","KS":"重复品名","PG":"测试"},
}
# Fixed standard algorithms, visible on maintenance page, not an arbitrary script engine.
RULES = [
    ("glue", "四位胶系", "型号2位 + 配方1位 + 功能用途1位；必须命中启用的主表记录，不从旧码默认推定。"),
    ("sheet", "片料尺寸", "英寸，单边精确到2位小数 ×100 后补足4位；两边拼为8位。不自动舍入。"),
    ("board", "基板厚度", "毫米 ×1000，补足4位；超出4位的大厚度定义待确认，禁止猜测。"),
    ("pp", "PP厚度 / RC", "厚度以μm整数编码（<300，末位0/3/5/8）；RC百分数×10（编码410至999）。未明确区间阻止生成。"),
    ("roll", "卷料尺寸", "严格使用卷料尺寸主表，R000XXXX 中的 X 是字面编码，不是待填写占位符。"),
]
