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
            context=text+' '+extra
            formulas=set(re.findall(r'配方\s*[:：]?\s*([A-Z])(?![A-Z])',context))
            if 'UL配方' in context:
                formulas.add('U')
            formula=next(iter(formulas)) if len(formulas)==1 else 'N' if not formulas else None
            uses={code for pattern,code in [(r'汽车板','A'),(r'MINI\s*LED','M'),
                  (r'LOW\s*DK','L'),(r'石英','Q'),(r'UL\s*认证样品','U')]
                  if re.search(pattern,context)}
            use=next(iter(uses)) if len(uses)==1 else 'N' if not uses else None
            if formula is None or use is None:
                warnings.append('配方或用途要求冲突，请核对胶系。')
            unique("glue", [r["code"] for r in candidates if formula and use and r["code"][2:] == formula+use],
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
        pair = re.search(r"(?<![A-Z0-9.])(H|\d+(?:\.\d+)?)\s*(?:OZ)?\s*/\s*(H|\d+(?:\.\d+)?)(?:\s*OZ)?(?![A-Z0-9.])", text)
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
            if pair[1] == '1' and pair[2] == 'H' and copper == 'G':
                top,bottom='H','1'  # Confirmed example only; do not infer general sorting.
            elif top != bottom:
                warnings.append('两面铜重不同，排列顺序尚待确认，请核对后选择铜箔组合。')
                top=bottom=''
            code = top+bottom+copper*2
            if any(r["category"] == "copper_pair" and r["code"] == code for r in active):
                put("copper_pair",code,"双面铜箔识别")
                values["copper_mode"]="standard"
            else:
                values.update(copper_mode="split", copper_top_weight=top, copper_bottom_weight=bottom, copper_top_type=copper, copper_bottom_type=copper)
        elif "copper_pair" in values:
            values["copper_mode"]="standard"
    else:
        numeric("pp_value", r"(?<![\d.])(\d+(?:\.\d+)?)\s*%")
        if "pp_value" not in values:
            numeric("pp_value", r"RC\s*(\d+(?:\.\d+)?)\s*%?")
        if "pp_value" in values:
            values["pp_mode"]="rc"
        else:
            numeric("pp_value", r"(?<![\d.])(\d+(?:\.\d+)?)\s*(?:UM|微米)\b")
            values["pp_mode"]="thickness"
        if "roll_size" in values:
            values["size_mode"]="roll"
        combined=text+' '+extra
        if re.search(r'卷|幅宽|(?<![A-Z])\d+\s*M(?![A-Z])', combined):
            values['size_mode']='roll'
            widths=re.findall(r'(?:幅宽|宽度)\s*[:：]?\s*(\d+(?:\.\d+)?)',combined)
            if not widths:
                widths=re.findall(r'(\d+(?:\.\d+)?)\s*(?:英寸|INCH|["”])',combined)
            width_codes={'49.7':'4970','43.8':'4380','49':'XXXX'}
            widths={str(Decimal(w).normalize()) for w in widths}
            lengths=set(re.findall(r'(?<![\d.])(\d+)\s*(?:M(?![A-Z])|米)',combined))
            if len(widths)==1 and next(iter(widths)) in width_codes and len(lengths)<=1 and '羽边' not in combined:
                length=next(iter(lengths),'0')
                if int(length)<=999:
                    put('roll_size','R'+length.zfill(3)+width_codes[next(iter(widths))], '卷长与幅宽识别；未注明卷长使用000')
                else:
                    values.pop('roll_size',None)
                    warnings.append('卷长超过3位编码范围，请核对。')
            elif 'roll_size' not in values:
                warnings.append('请核对卷料幅宽；特殊幅宽及羽边规则尚待确认，不自动选择。')
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
