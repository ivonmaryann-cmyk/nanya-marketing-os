"""Deterministic new-name recognition. No legacy codes or remote model calls."""
import re
from decimal import Decimal

# Document examples are explicit, reviewable defaults, never database ordering.
DEFAULTS = {"board_spec": "C3", "structure": "A", "glass_type": "C",
            "pp_spec": "XX", "marking": "W", "board_grade": "A0", "pp_grade": "A0"}


def normalize(text):
    return str(text).upper().replace("－", "-").replace("μ", "U").replace("Μ", "U")


def extract(product, specification, requirements, rows):
    text = normalize(specification)
    extra = normalize(requirements)
    values, evidence, warnings = {"size_mode": "sheet"}, {}, []
    active = [r for r in rows if r["enabled"]]

    def put(key, value, source):
        values[key] = value
        evidence[key] = source

    def unique(key, candidates, source):
        candidates = set(candidates)
        if len(candidates) == 1:
            put(key, candidates.pop(), source)
        elif candidates:
            values.pop(key, None)
            evidence.pop(key, None)
            warnings.append(f"{key} 存在多个候选，请核对")
        return bool(candidates) or key in values

    for category in {r["category"] for r in active}:
        if category == "glue":
            continue
        subset = [r for r in active if r["category"] == category]
        hits = []
        for fragment in (text, extra):
            matches = []
            for r in subset:
                aliases = [x.strip() for x in r["details"].get("aliases", "").split("|") if x.strip()]
                # Names are not automatically aliases: e.g. "三级" cannot decide C3 vs T3.
                aliases += [r["code"]] if len(r["code"]) >= 2 else []
                for alias in aliases:
                    if re.search(r"(?<![A-Z0-9])" + re.escape(normalize(alias)) + r"(?![A-Z0-9])", fragment):
                        matches.append((len(alias), r["code"]))
            if matches:
                longest = max(n for n, _ in matches)
                hits += [code for n, code in matches if n == longest]
        if hits:
            unique(category, hits, "规格 / 特殊要求识别")
        else:
            defaults = [r["code"] for r in subset if product in r["details"].get(
                "default_for", "board,pp" if DEFAULTS.get(category) == r["code"] else "").split(",")]
            unique(category, defaults, "规则默认 · 请确认")

    glue_rows = [r for r in active if r["category"] == "glue"]
    exact = [r["code"] for r in glue_rows if re.search(r"(?<![A-Z0-9])"+r["code"]+r"(?![A-Z0-9])", text+" "+extra)]
    if exact:
        unique("glue", exact, "四位胶系识别")
    else:
        pattern = r"NY\s*-?\s*[A-Z]?\d+[A-Z]*(?:\([A-Z]\))?"
        models = [m.removesuffix("P") for m in re.findall(pattern, text+" "+extra)]
        model = re.sub(r"\s", "", models[0]) if len(set(models)) == 1 else None
        candidates = []
        if model:
            for r in glue_rows:
                names = re.findall(pattern, normalize(r["name"]))
                if model in [re.sub(r"\s", "", n) for n in names]:
                    candidates.append(r)
            formula = "U" if "UL配方" in text+extra else "N"
            use = "A" if "汽车板" in text+extra else "M" if re.search(r"MINI\s*LED", text+extra) else "N"
            unique("glue", [r["code"] for r in candidates if r["code"][2:] == formula+use],
                   "型号识别；未注明配方 / 用途使用通用项，请确认")
        aliases = [r["code"] for r in glue_rows for a in r["details"].get("aliases", "").split("|") if a.strip() and normalize(a.strip()) in text+" "+extra]
        if aliases:
            unique("glue", aliases + ([values['glue']] if 'glue' in values else []), "胶系别名映射")

    def numeric(key, pattern, source=None, scale=1):
        source = text+" "+extra if source is None else source
        hits = re.findall(pattern, source)
        if hits:
            unique(key, [str(Decimal(v)*Decimal(str(scale))) for v in hits], "规格提取")

    if product == "board":
        numeric("thickness", r"(?<![\d.])(\d+(?:\.\d+)?)\s*MM\b")
        if "thickness" not in values:
            numeric("thickness", r"(?<![\d.])(\d+(?:\.\d+)?)\s*MIL\b", scale="0.0254")
        pair = re.search(r"(?<![A-Z0-9.])(H|\d+(?:\.\d+)?)\s*/\s*(H|\d+(?:\.\d+)?)(?![A-Z0-9.])", text)
        types = [("RTF6","U"),("RTF5","T"),("RTF4","M"),("RTF3","G"),("RTF2","F"),("RTF","R"),("HTE","N")]
        # Chinese text is a Unicode word character; \b incorrectly misses HTE铜箔.
        type_hits = {code for alias,code in types if re.search(r"(?<![A-Z0-9])"+alias+r"(?![A-Z0-9])", text)}
        if values.get("copper_type"):
            type_hits.add(values["copper_type"])
        copper = next(iter(type_hits)) if len(type_hits)==1 else None
        if len(type_hits)>1:
            warnings.append("检测到不同铜箔类型，请分别核对上下两面")
        single_h = not pair and re.search(r"(?<![A-Z0-9])H\s*OZ(?![A-Z0-9])", text)
        if single_h and copper == 'N' and any(r['category']=='copper_pair' and r['code']=='HHNN' for r in active):
            put('copper_pair','HHNN','按文档双面H铜箔示例自动填写 · 请确认')
            values['copper_mode']='standard'
            warnings.append('Hoz HTE已按双面H铜箔填写，请确认客户是否有单面或不同铜重要求。')
        elif pair and copper:
            def weight(raw):
                if raw == 'H':
                    return 'H'  # Only usable through a documented complete pair unless maintained.
                matches=[r['code'] for r in active if r['category']=='copper_weight' and
                         (raw==r['name'].split(' ')[0] or raw in r['details'].get('aliases','').split('|'))]
                return matches[0] if len(matches)==1 else ''
            top,bottom=weight(pair[1]),weight(pair[2])
            code = top+bottom+copper*2
            if any(r["category"] == "copper_pair" and r["code"] == code for r in active):
                put("copper_pair",code,"双面铜箔识别")
                values["copper_mode"]="standard"
            else:
                values.update(copper_mode="split", copper_top_weight=top, copper_bottom_weight=bottom, copper_top_type=copper, copper_bottom_type=copper)
        elif "copper_pair" in values:
            values["copper_mode"]="standard"
    else:
        numeric("pp_value", r"RC\s*(\d+(?:\.\d+)?)\s*%?")
        if "pp_value" in values:
            values["pp_mode"]="rc"
        else:
            numeric("pp_value", r"(?<![\d.])(\d+(?:\.\d+)?)\s*(?:UM|微米)\b")
            values["pp_mode"]="thickness"
        if "roll_size" in values:
            values["size_mode"]="roll"
    sizes = re.findall(r'(\d+(?:\.\d+)?)\s*["”]?\s*[X×*]\s*(\d+(?:\.\d+)?)\s*(["”]|INCH|英寸|MM)?', text)
    # Exclude layups such as 1x2116; bare sheet dimensions must both be >= 10.
    sizes = [(a,b,u) for a,b,u in sizes if u or (Decimal(a)>=10 and Decimal(b)>=10)]
    if len(sizes)==1:
        a,b,unit=sizes[0]
        divisor=Decimal("25.4") if unit=="MM" else Decimal(1)
        put("length",str(Decimal(a)/divisor),"尺寸提取（英寸）")
        put("width",str(Decimal(b)/divisor),"尺寸提取（英寸）")
    elif sizes:
        warnings.append("存在多组尺寸，请选择本次转换的尺寸")
    combined=text+" "+extra
    without_negative=combined.replace("不印字", "").replace("无水印", "")
    if ("无水印" in combined or "不印字" in combined) and "印字" in without_negative:
        values.pop('marking',None)
        warnings.append("印字要求冲突，请确认使用印字还是无水印")
    elif "无水印" in combined or "不印字" in combined:
        put("marking","W","规格 / 特殊要求识别")
    elif "印字" in combined:
        put("marking","Q","规格 / 特殊要求识别")
    if extra:
        warnings.append("特殊要求已参与基础识别，请核对未结构化要求；本期不启用客户专属规则。")
    return {"values":values,"evidence":evidence,"warnings":warnings}
